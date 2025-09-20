import sys
import xml.etree.ElementTree as ET

try:
    from win32com.client import Dispatch, constants  # poetry install pywin32 (32-bit)
except ImportError:
    print("Please 'poetry add pywin32' (use 32-bit Python for QB Desktop).")
    sys.exit(1)


def build_terms_query() -> str:
    """Return a minimal TermsQueryRq XML."""
    return """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <TermsQueryRq/>
  </QBXMLMsgsRq>
</QBXML>"""


def parse_and_print(response_xml: str) -> None:
    """Parse response and print term name + discount days."""
    root = ET.fromstring(response_xml)

    terms_query_rs = root.find(".//TermsQueryRs")
    if terms_query_rs is None:
        print("No TermsQueryRs in response")
        return

    # Standard terms
    for term in terms_query_rs.findall("StandardTermsRet"):
        name = term.findtext("Name", default="(no name)")
        discount_days = term.findtext("StdDiscountDays", default="N/A")
        print(f"Name: {name}, Discount Days: {discount_days}")

    # Date-driven terms
    for term in terms_query_rs.findall("DateDrivenTermsRet"):
        name = term.findtext("Name", default="(no name)")
        day_due = term.findtext("DayOfMonthDue", default="N/A")
        print(f"Name: {name}, Day Of Month Due: {day_due}")


def main():
    ct_local_qbd = getattr(constants, "ctLocalQBD", 1)
    om_dont_care = getattr(constants, "omDontCare", 0)

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
