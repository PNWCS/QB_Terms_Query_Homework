import sys
import xml.etree.ElementTree as ET

try:
    from win32com.client import Dispatch, constants  # poetry install pywin32 (32-bit)
except ImportError:
    print("Please 'poetry add pywin32' (use 32-bit Python for QB Desktop).")
    sys.exit(1)


def build_terms_query() -> str:
    """Return a minimal TermsQueryRq QBXML request as a string."""
    return """<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <TermsQueryRq requestID="1">
      <!-- Optional filters could go here, e.g., ListID, FullName, ActiveStatus -->
    </TermsQueryRq>
  </QBXMLMsgsRq>
</QBXML>
"""


def parse_and_print(response_xml: str) -> None:
    """Parse response and print term name + discount days."""
    root = ET.fromstring(response_xml)
    rs = root.find(".//TermsQueryRs")

    if rs is None:
        print("No TermsQueryRs found in response")
        return
    status_code = int(rs.attrib.get("statuscode", "0"))
    status_msg = rs.attrib.get("statusMessage", "")
    if status_code != 0:
        print(f"QuickBook returned error {status_code}:{status_msg}")
        return
    for std in rs.findall(".//StandardTermsRet"):
        name = (std.findtext("Name") or "").strip()
        discount_days = (std.findtext("StdDiscountDays") or "").strip()
        print(f"Term: {name}, Discount Days: {discount_days}")


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
