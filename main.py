import os
import comtypes.client


class MSOfficeToPDFConverter:
    _app = None
    _format = None

    file_formats = {"word": 17, "excel": 57, "powerpoint": 32}

    def __init__(self) -> None:
        pass

    def convert_to_pdf(self):
        directory = input(
            "Enter the path of the directory containing Microsoft Office files: "
        )

        self._init_conversion(directory)

    def _init_conversion(self, directory):
        for root, _, files in os.walk(directory):
            self._convert_files_in_dir_to_pdf(files=files, root_folder=root)

    def _convert_files_in_dir_to_pdf(self, files, root_folder):
        for file in files:
            input_file = os.path.join(root_folder, file)
            filename, extension = os.path.splitext(input_file)

            if self._check_if_extension_is_valid(extension=extension):
                self._convert_file_to_pdf(input_file, filename + ".pdf")

    def _check_if_extension_is_valid(self, extension):
        valid_extensions = [".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"]
        return extension.lower() in valid_extensions

    def _convert_file_to_pdf(self, input_file, output_file):
        initialized_file = None
        file_formats = {"word": 17, "excel": 57, "powerpoint": 32}

        try:
            # Check the file extension to determine the application to use
            _, extension = os.path.splitext(input_file)
            extension = extension.lower()

            if extension == ".docx" or extension == ".doc":
                self._app = comtypes.client.CreateObject("Word.Application")
                self._format = file_formats["word"]
                initialized_file = self._app.Documents.Open(input_file)

            elif extension == ".xlsx" or extension == ".xls":
                self._app = comtypes.client.CreateObject("Excel.Application")
                self._format = file_formats["excel"]
                initialized_file = self._app.Workbooks.Open(input_file)

            elif extension == ".pptx" or extension == ".ppt":
                self._app = comtypes.client.CreateObject("PowerPoint.Application")
                self._format = file_formats["powerpoint"]
                initialized_file = self._app.Presentations.Open(input_file)

            else:
                raise Exception(f"Unsupported file format: {input_file}")

            initialized_file.SaveAs(output_file, FileFormat=self._format)
            initialized_file.Close()
            print(f"Converted: {input_file} -> {output_file}")

        except Exception as e:
            print(f"Error converting file: {str(e)}")

        finally:
            # Close all applications
            self._app.Quit()


if __name__ == "__main__":
    converter = MSOfficeToPDFConverter()
    converter.convert_to_pdf()
