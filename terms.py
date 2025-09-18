import sys
import xml.etree.ElementTree as ET

try:
    from win32com.client import (
        Dispatch,
        constants,
    )  # Requires 32-bit Python for QB Desktop
except ImportError:
    print("Please 'poetry add pywin32' (use 32-bit Python for QB Desktop).")
    sys.exit(1)


def build_terms_query() -> str:
    """Return a minimal TermsQueryRq XML."""
    qbxml = """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="16.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <TermsQueryRq requestID="1">
      <ActiveStatus>All</ActiveStatus>
    </TermsQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""
    return qbxml


def parse_and_print(response_xml: str) -> None:
    """Parse response and print term name + discount information."""
    try:
        root = ET.fromstring(response_xml)
    except ET.ParseError as e:
        print(f"Error parsing response XML: {e}")
        return

    # Handle both StandardTermsRet and DateDrivenTermsRet
    standard_terms = root.findall(".//StandardTermsRet")
    date_driven_terms = root.findall(".//DateDrivenTermsRet")

    if not (standard_terms or date_driven_terms):
        print("No terms found in response.")
        return

    # Standard terms
    for term in standard_terms:
        name = term.findtext("Name", default="(unknown)")
        discount_days = term.findtext("StdDiscountDays", default="N/A")
        discount_pct = term.findtext("DiscountPct", default="N/A")
        print(
            f"[Standard] Term: {name}, Discount Days: {discount_days}, Discount %: {discount_pct}"
        )

    # Date-driven terms
    for term in date_driven_terms:
        name = term.findtext("Name", default="(unknown)")
        discount_day = term.findtext("DiscountDayOfMonth", default="N/A")
        discount_pct = term.findtext("DiscountPct", default="N/A")
        print(
            f"[DateDriven] Term: {name}, Discount Day: {discount_day}, Discount %: {discount_pct}"
        )


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

        # Save to file for debugging
        with open("response.xml", "w", encoding="utf-8") as file:
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
