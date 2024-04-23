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
from create_word_doc import CreateWordDoc


def main(page: Page):
    page.title = "Bill Creation Tool"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.bgcolor = ft.colors.INDIGO_300
    page.window_width = 800
    page.window_height = 700

    selected_file_name = Text()
    selected_path = Text()
    output_path = Text()
    txt_name = TextField(hint_text="Enter file name here", border="none")
    welcome_text = Text("Welcome to Bill Creation Tool:\n "
                        "- just pick a file\n"
                        "- then choose the output folder\n"
                        "- enter the name of the bill\n "
                        "- and finally hit Create Bill",
                        size=20,
                        color=ft.colors.WHITE,
                        text_align=ft.TextAlign.JUSTIFY,
                        font_family="Courier New")

    def pick_input_file(e: FilePickerResultEvent):
        if e.files:
            selected_file_name.value = e.files[0].name
            selected_file_name.update()
            selected_path.value = e.files[0].path

    pick_files_dialog = FilePicker(on_result=pick_input_file)

    def get_output_directory(e: FilePickerResultEvent):
        output_path.value = e.path
        output_path.update()

    get_directory_dialog = FilePicker(on_result=get_output_directory)

    def output_filename(e: TextField):
        output_file_name = txt_name.value
        input_pdf_path = selected_path.value
        output_doc_path = output_path.value

        word_doc = CreateWordDoc(input_pdf_path, output_doc_path, output_file_name)
        word_doc.create_bill()

    page.overlay.extend([pick_files_dialog, get_directory_dialog])

    page.add(
        ft.Container(
            content=welcome_text,
            alignment=ft.alignment.center,
            margin=20,
            padding=10
        ),
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
                    bgcolor=ft.colors.WHITE,
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
                    bgcolor=ft.colors.WHITE,
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
                    bgcolor=ft.colors.WHITE,
                    padding=10,
                    width=500,
                    height=50,
                    border_radius=10
                ),
            ],
        ),
    )


ft.app(target=main)

