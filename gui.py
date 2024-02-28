import flet as ft
from flet import (
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    Row,
    Text,
    icons,
)


def main(page: Page):
    page.title = "Bill Creation Tool"

    selected_file_name = Text()
    selected_path = Text()

    def pick_files_result(e: FilePickerResultEvent):
        if e.files:
            selected_file_name.value = e.files[0].name
            selected_file_name.update()
            selected_path.value =  e.files[0].path
            print(selected_path.value)

    pick_files_dialog = FilePicker(on_result=pick_files_result)

    page.overlay.extend([pick_files_dialog])

    page.add(
        Row(
            [
                ft.Container(
                    ElevatedButton(
                        "Pick files",
                        icon=icons.UPLOAD_FILE,
                        on_click=lambda _: pick_files_dialog.pick_files()
                    ),
                    margin=10,
                    padding=10,
                    alignment=ft.alignment.center,
                    bgcolor=ft.colors.GREEN_100,
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
            alignment = ft.MainAxisAlignment.CENTER,
        ),
    )


ft.app(target=main)
