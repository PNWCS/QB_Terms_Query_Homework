import sys
import xml.etree.ElementTree as ET

try:
    from win32com.client import Dispatch, constants  # poetry install pywin32 (32-bit)
except ImportError:
    print("Please 'poetry add pywin32' (use 32-bit Python for QB Desktop).")
    sys.exit(1)


def build_terms_query() -> str:
    """Return a minimal TermsQueryRq XML (only required fields)."""
    return """<?xml version="1.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <TermsQueryRq requestID="1" />
  </QBXMLMsgsRq>
</QBXML>"""


def parse_and_print(response_xml: str) -> None:
    """Parse response and print term name + discount days."""
    root = ET.fromstring(response_xml)

    # Get the statusCode and statusMessage from the response
    msgs_response = root.find(".//QBXMLMsgsRs/TermsQueryRs")
    if msgs_response is not None:
        status_code = msgs_response.get("statusCode")
        status_message = msgs_response.get("statusMessage")

        if status_code != "0":  # Not successful
            print(f"Error {status_code}: {status_message}")
            return

    # Loop through all StandardTermsRet tags
    for terms in root.findall(".//StandardTermsRet"):
        name = terms.find("Name")
        discount_days = terms.find("DiscountDays")
        print(f"{name.text if name is not None else 'N/A'} - {discount_days.text if discount_days is not None else 'N/A'}")


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