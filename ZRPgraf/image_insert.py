from openpyxl.drawing.image import Image
from openpyxl import load_workbook


def image_insert(excel_name: str, image: str, worksheet_name: str) -> None:
    """
    Функция добавления изображения в указанный excel документ
    :param excel_name: Название документа excel
    :param image: Название изображения
    :param worksheet_name: Наименование листа в файле excel
    :return: None
    """
    wb = load_workbook(excel_name)
    ws3 = wb[worksheet_name]
    image = Image(image)
    image.height = 13 * 8 * 4
    image.width = 13 * 40 * 4

    ws3.add_image(image, "M2")
    wb.save(excel_name)
