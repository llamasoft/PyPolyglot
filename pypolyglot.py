import argparse
import io
import logging
import pathlib
import shlex
import shutil
import sys
import typing
import xml.etree.ElementTree as ET
import zipfile


logging.basicConfig(
    format="[%(levelname)s] %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)


class PyPoly:
    OFFICE_XML = "[Content_Types].xml"

    @classmethod
    def polyglotify(cls, input_path: pathlib.Path, payload_path: pathlib.Path, output_path: pathlib.Path) -> None:
        """Creates a modified version of the input file that is also an Python archive."""

        if not input_path.exists():
            raise RuntimeError(f"Input path {input_path} does not exist.")

        if not payload_path.exists():
            raise RuntimeError(f"Payload script path {payload_path} does not exist.")

        if output_path.exists() and output_path.samefile(input_path):
            raise RuntimeError("Input and output paths cannot be the same.")

        with open(payload_path) as f:
            payload = f.read()

        # If the input isn't a zip file, just add a custom zip file to the end.
        # The zip file central directory structure is stored at the end of the file
        # and this is actually what most tools (Python included) look for.
        if not cls.is_zip(input_path):
            return cls.plain_to_pyarchive(input_path, payload, output_path)

        # The input file is a zip file of some sort!
        with zipfile.ZipFile(input_path, "r") as input_zf:
            if cls.is_pyarchive(input_zf):
                # Nothing for us to do, the file has already been treated.
                raise RuntimeError("Input path is already an Python archive.")

            if not cls.is_office_doc(input_zf):
                return cls.zip_to_pyarchive(input_path, payload, output_path)

            return cls.office_to_pyarchive(input_path, payload, output_path)

    # Unfortunately as of Oct 2025, Python's zipfile library doesn't support removing/overwriting existing files:
    #   https://github.com/python/cpython/issues/51067
    # Instead, we must copy everything except the file that we want to remove/replace.
    @staticmethod
    def copy_and_filter_zip(
        input_zf: zipfile.ZipFile,
        output_zf: zipfile.ZipFile,
        filter_fn: typing.Callable[[zipfile.ZipInfo],bool]
    ) -> None:
        """Copies the filtered contents of one ZipFile to another."""
        for file_info in input_zf.infolist():
            if filter_fn(file_info):
                log.debug(f"Skipping {file_info.filename}")
                continue
            file_data = input_zf.read(file_info)
            output_zf.writestr(file_info, file_data)

    @staticmethod
    def is_zip(file_or_path: str) -> bool:
        return zipfile.is_zipfile(file_or_path)

    @classmethod
    def is_pyarchive(cls, zf: zipfile.ZipFile) -> bool:
        """Returns if the given zip file is already a Python archive."""
        try:
            zf.getinfo("__main__.py")
            log.debug("Found an existing __main__.py")
            return True
        except KeyError:
            return False

    @classmethod
    def is_office_doc(cls, zf: zipfile.ZipFile) -> bool:
        """Returns if the given zip file is a Microsoft Office document."""
        try:
            zf.getinfo(cls.OFFICE_XML)
            log.debug(f"Found an existing {cls.OFFICE_XML}")
            return True
        except KeyError:
            return False

    @classmethod
    def create_pyarchive(cls, payload: str) -> bytes:
        """Creates an Python archive and returns it as bytes."""
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("__main__.py", payload, compress_type=zipfile.ZIP_DEFLATED)
        return buf.getvalue()

    @classmethod
    def plain_to_pyarchive(cls, input_path: pathlib.Path, payload: str, output_path: pathlib.Path) -> None:
        """Creates an Python archive by simply appending a formatted zip archive."""
        log.debug("Input path appears to be a generic file")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(input_path, output_path)
        with open(output_path, "ab") as f:
            log.info("Creating Python script payload")
            f.write(cls.create_pyarchive(payload))

    @classmethod
    def zip_to_pyarchive(cls, input_path: pathlib.Path, payload: str, output_path: pathlib.Path) -> None:
        """Creates an Python archive by adding a __main__.py to an existing zip archive."""
        log.debug("Input path appears to be a zip archive")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(input_path, output_path)
        with zipfile.ZipFile(output_path, "a") as output_zf:
            log.info("Adding Python script payload")
            output_zf.writestr("__main__.py", payload, compress_type=zipfile.ZIP_DEFLATED)

    @classmethod
    def office_to_pyarchive(cls, input_path: pathlib.Path, payload: str, output_path: pathlib.Path) -> None:
        """Creates an Python archive by modifying a Microsoft Office document zip file."""
        log.debug("Input path appears to be an Office document")
        with zipfile.ZipFile(input_path, "r") as input_zf, zipfile.ZipFile(output_path, "w") as output_zf:
            # Copy everything but the content type XML file because zipfile doesn't support
            # replacing files within an existing archive.  See copy_and_filter_zip() for details.
            log.info("Copying document contents to new archive")
            cls.copy_and_filter_zip(input_zf, output_zf, filter_fn=lambda info: info.filename == cls.OFFICE_XML)

            # Add the Python script payload.
            log.info("Adding Python script payload")
            output_zf.writestr("__main__.py", payload, compress_type=zipfile.ZIP_DEFLATED)

            # Extract the source's existing content type XML file for patching.
            # This prevents Microsoft Office from throwing errors.
            # See patch_office_xml() function for details.
            log.info(f"Patching {cls.OFFICE_XML} file")
            content_types_xml = input_zf.read(cls.OFFICE_XML)
            patched_types_xml = cls.patch_office_xml(content_types_xml)
            output_zf.writestr(cls.OFFICE_XML, patched_types_xml, compress_type=zipfile.ZIP_DEFLATED)

    @classmethod
    def patch_office_xml(cls, content_types_xml: bytes, archive_paths: list = None) -> bytes:
        """Returns a patched [Content_Types].xml file contents to prevent Microsoft Office warnings."""
        if archive_paths is None:
            archive_paths = ["__main__.py"]

        # The zip archive's paths don't start with a slash, but they do in the XML file entries.
        # Normalize the archive path list to match what the XML file will contain.
        archive_paths = ["/" + path.lstrip("/") for path in archive_paths]

        # Microsoft Office is *very* particular about its document formats.
        # If it detects any oddities, it will throw a "found unreadable content" error.
        # Things that will cause the check to fail:
        # - Any additional data before or after the Office document zip file.
        # - Any additional unreferenced files within the Office document zip file.
        # The only portable solution that works for all Office products is to
        # add references to our extra files to the [Content_Types].xml file.

        # Microsoft Office is also pretty picky about the [Content_Types].xml file format.
        # It's just XML, and Python's ElementTree library is more than happy to parse it,
        # but unless we set the proper global schema namespace, the patched file is going
        # to include an explicit namespace prefix on every element.
        root = ET.fromstring(content_types_xml)

        # Get the default namespace from the root element's tag (i.e. {URL}TagName -> URL)
        namespace = root.tag[1 : root.tag.index("}")]

        # Yes, the namespace registry is global. Yes, this feels dirty.
        # ElementTree.tostring has a default_namespace argument, but it only applies
        # to non-namespaced tags.  If the tree has any namespaced tags (which we do),
        # setting default_namespace will cause it to throw an exception.
        ET.register_namespace("", namespace)

        # Add any missing overrides to the XML with ContentType="plain/text".
        existing_overrides = set(
            elem.attrib["PartName"]
            for elem in root.findall("{*}Override")
            if "PartName" in elem.attrib
        )
        log.debug(f"{cls.OFFICE_XML} contains {len(existing_overrides)} entries: {existing_overrides}")

        missing_overrides = set(archive_paths) - existing_overrides
        log.debug(f"Adding {len(missing_overrides)} new entries: {missing_overrides}")

        for path in missing_overrides:
            override = ET.SubElement(root, "{" + namespace + "}Override")
            override.attrib["PartName"] = path
            override.attrib["ContentType"] = "plain/text"

        # Return the patched XML with an XML declaration and UTF-8 encoding.
        return ET.tostring(root, encoding="unicode", xml_declaration=True)


def main():
    # Include the full help listing (--help) instead of short usage message on errors.
    # This is useful for when this script is itself is a polyglot payload and the user
    # is unaware that the script accepts parameters.
    class HelpfulParser(argparse.ArgumentParser):
        def format_usage(self, *args, **kwargs):
            return self.format_help(*args, **kwargs)

    this_file = sys.argv[0]
    easter_egg = (
        "I hope you didn't blindly run this polyglot!\n"
        f"You can extract this script with: unzip {shlex.quote(this_file)} __main__.py\n"
        "For info on how this works and how it was created, see:\n"
        "  https://github.com/llamasoft/PyPolyglot\n"
    )
    parser = HelpfulParser(
        description=(
            "Creates polyglots by adding a Python script payload to a file.\n"
            "Includes special support for Microsoft Office documents."
        ),
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=None if this_file.endswith(".py") else easter_egg,
    )
    parser.add_argument(
        "input_path",
        metavar="INPUT",
        type=pathlib.Path,
        help="Input file to convert into a Python archive polyglot",
    )
    parser.add_argument(
        "payload_path",
        metavar="PAYLOAD",
        type=pathlib.Path,
        help="Path to a Python script to use as the archive's payload",
    )
    parser.add_argument(
        "output_path",
        metavar="OUTPUT",
        type=pathlib.Path,
        help="Output file for the Python archive polyglot",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args()

    if args.verbose:
        log.setLevel(logging.DEBUG)

    PyPoly.polyglotify(args.input_path, args.payload_path, args.output_path)

    pyexe = pathlib.Path(sys.executable).name
    quoted_path = shlex.quote(str(args.output_path))
    print("Python archive polyglot created!")
    print("You can test the archive using:")
    print(f"  {pyexe} {quoted_path}")


if __name__ == "__main__":
    main()
