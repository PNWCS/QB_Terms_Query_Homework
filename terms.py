import sys
import xml.etree.ElementTree as ET

try:
    from win32com.client import Dispatch, constants  # poetry install pywin32 (32-bit)
except ImportError:
    print("Please 'poetry add pywin32' (use 32-bit Python for QB Desktop).")
    sys.exit(1)


def build_terms_query() -> str:
    """Return a minimal TermsQueryRq XML."""

    xml_string = """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <TermsQueryRq/>
  </QBXMLMsgsRq>
</QBXML>

     """

    return xml_string


def parse_and_print(response_xml: str) -> None:
    """Parse response and print term name + discount days."""
    with open("response.xml", "w") as file:
        file.write(response_xml)

    tree = ET.parse("response.xml")
    root = tree.getroot()

    terms_query_rs = root.find(".//TermsQueryRs")

    status_code = terms_query_rs.get("statusCode")
    status_message = terms_query_rs.get("statusMessage")

    if status_code != "0":
        print(f"Error: {status_message}")
        exit()

    for term in terms_query_rs.findall("StandardTermsRet"):
        name = term.findtext("Name")
        discount_days = term.findtext("StdDiscountDays")
        print(f"Name: {name}, Discount Days: {discount_days}")


def main():
    # Fallbacks if pywin32 doesn't expose these constants on your machine
    ct_local_qbd = getattr(constants, "ctLocalQBD", 1)  # 1 = local QBD
    om_dont_care = getattr(constants, "omDontCare", 0)  # 0 = DoNotCare

    rp = None
    ticket = None
    try:
        rp = Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection2("", "Python QBXML Demo", ct_local_qbd)
        ticket = rp.BeginSession("", om_dont_care)

        request_xml = build_terms_query()
        response_xml = rp.ProcessRequest(ticket, request_xml)
        with open("response.xml", "w") as file:
            file.write(response_xml)
        parse_and_print(response_xml)

    finally:
        if rp and ticket:
            try:
                rp.EndSession(ticket)
            except Exception:
                pass
        if rp:
            try:
                rp.CloseConnection()
            except Exception:
                pass


if __name__ == "__main__":
    main()
