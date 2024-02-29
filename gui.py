import flet as ft
from flet import (
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    Row,
    Text,
    TextField,
    icons,
)
from createWordDoc import CreateWordDoc

def main(page: Page):
    page.title = "Bill Creation Tool"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    selected_file_name = Text()
    selected_path = Text()
    output_path = Text()
    txt_name = TextField(label="Your file name")

    def pick_input_file(e: FilePickerResultEvent):
        if e.files:
            selected_file_name.value = e.files[0].name
            selected_file_name.update()
            selected_path.value =  e.files[0].path

    pick_files_dialog = FilePicker(on_result=pick_input_file)

    def get_output_directory(e: FilePickerResultEvent):
        output_path.value = e.path
        #print("####Output directory: ", output_path.value)
        output_path.update()

    get_directory_dialog = FilePicker(on_result=get_output_directory)

    def output_filename(e: TextField):
        if not txt_name.value:
            txt_name.error_text = "Please enter your file name"
            page.update()
        else:
            output_file_name = txt_name.value
            input_doc_path = selected_path.value
            output_doc_path = output_path.value
            print("output_doc_path", output_doc_path)
            print("input_doc_path", input_doc_path)
            print(output_file_name)
            word_doc = CreateWordDoc(input_doc_path, output_doc_path, output_file_name)
            word_doc.create_bill()

    page.overlay.extend([pick_files_dialog, get_directory_dialog])

    page.add(
        Row(
            [
                ft.Container(
                    ElevatedButton(
                        "Pick file",
                        icon=icons.UPLOAD_FILE,
                        on_click=lambda _: pick_files_dialog.pick_files()
                    ),
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.GREEN_100,
                    margin=10,
                    padding=10,
                    width=250,
                    height=100,
                    border_radius=30,

                ),
                ft.Container(
                    content=selected_file_name,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.GREEN_50,
                    width=500,
                    height=50,
                    border_radius=10,
                ),

            ],
        ),
        Row(
            [
                ft.Container(
                    ElevatedButton(
                        "Output directory",
                        icon=icons.FOLDER_OPEN,
                        on_click=lambda _: get_directory_dialog.get_directory_path(),
                        disabled=page.web,
                    ),
                    margin=10,
                    padding=10,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.BLUE_100,
                    width=250,
                    height=100,
                    border_radius=30,
                ),
                ft.Container(
                    content=output_path,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.BLUE_50,
                    width=500,
                    height=50,
                    border_radius=10,
                ),
            ],
        ),
        Row(
            [
                ft.Container(
                    ElevatedButton(
                        "Create Bill",
                        icon=icons.CREATE_OUTLINED,
                        on_click=output_filename,
                        disabled=page.web,
                    ),
                    margin=10,
                    padding=10,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.INDIGO_100,
                    width=250,
                    height=100,
                    border_radius=30,
                ),
                ft.Container(
                    content=txt_name,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.INDIGO_50,
                    width=500,
                    height=50,
                ),
            ],
        ),
    )

ft.app(target=main)

