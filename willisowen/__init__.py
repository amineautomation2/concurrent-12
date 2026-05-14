from .investment import Willis_IF
from .etf import Willis_ETF
from .mutual_fund import Willis_MF
from utils import clean_spreadsheet, delay, get_xlsx_filepath, get_random_user_agent


def willis_owen_runner() -> None:
    output_xlsx = get_xlsx_filepath("willis_owen.xlsx")
    headers = get_random_user_agent()
    clean_spreadsheet(output_xlsx)

    Willis_IF(headers, output_xlsx)

    delay(10, 20)
    Willis_ETF(headers, output_xlsx)

    delay(10, 20)
    Willis_MF(headers, output_xlsx)
