"""
Tally HTTP Connector
Direct integration with local Tally Prime / Tally ERP 9 XML Server (default http://localhost:9000).
"""

import requests
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional


class TallyConnector:
    def __init__(self, host: str = "localhost", port: int = 9000, timeout: int = 15):
        self.host = host.strip()
        self.port = int(port)
        self.timeout = timeout

    @property
    def base_url(self) -> str:
        h = self.host
        if not h.startswith("http://") and not h.startswith("https://"):
            h = f"http://{h}"
        return f"{h}:{self.port}"

    def test_connection(self, timeout: Optional[int] = None) -> Dict[str, Any]:
        """
        Tests connection to Tally and retrieves active company list.
        """
        t = timeout if timeout is not None else self.timeout
        request_xml = """<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>List of Companies</REPORTNAME>
        <STATICVARIABLES>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        </STATICVARIABLES>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""
        try:
            resp = requests.post(
                self.base_url,
                data=request_xml.encode("utf-8"),
                headers={"Content-Type": "text/xml;charset=utf-8"},
                timeout=t
            )
            if resp.status_code == 200:
                companies = self._parse_companies(resp.text)
                return {
                    "connected": True,
                    "status_code": resp.status_code,
                    "companies": companies,
                    "active_company": companies[0] if companies else "Active (Name unlisted)",
                    "message": f"Successfully connected to Tally at {self.base_url}"
                }
            else:
                return {
                    "connected": False,
                    "status_code": resp.status_code,
                    "message": f"Tally responded with HTTP {resp.status_code}"
                }
        except requests.exceptions.ConnectionError:
            return {
                "connected": False,
                "status_code": None,
                "message": (
                    f"Could not connect to Tally at {self.base_url}.\n"
                    "Please ensure Tally is running and XML/ODBC Server is enabled.\n"
                    "(In Tally Prime: F1:Help > Settings > Connectivity > Client/Server configuration > "
                    "TallyPrime acts as: Both/Server, Port: 9000)"
                )
            }
        except requests.exceptions.Timeout:
            return {
                "connected": False,
                "status_code": None,
                "message": f"Connection to Tally timed out after {self.timeout}s."
            }
        except Exception as e:
            return {
                "connected": False,
                "status_code": None,
                "message": f"Connection error: {str(e)}"
            }

    def _parse_companies(self, xml_text: str) -> List[str]:
        companies = []
        try:
            root = ET.fromstring(xml_text)
            for elem in root.iter():
                tag_upper = elem.tag.upper()
                if "COMPANY" in tag_upper and elem.text and elem.text.strip():
                    txt = elem.text.strip()
                    if txt not in companies and not txt.startswith("$$"):
                        companies.append(txt)
        except Exception:
            pass
        return companies

    def send_xml(self, xml_payload: str) -> Dict[str, Any]:
        """
        Sends an import XML payload to Tally and parses the resulting <RESPONSE>.
        """
        try:
            resp = requests.post(
                self.base_url,
                data=xml_payload.encode("utf-8"),
                headers={"Content-Type": "text/xml;charset=utf-8"},
                timeout=self.timeout
            )
            if resp.status_code == 200:
                return self._parse_import_response(resp.text)
            else:
                return {
                    "success": False,
                    "created": 0,
                    "altered": 0,
                    "errors": 1,
                    "ignored": 0,
                    "messages": [f"HTTP Error {resp.status_code}: {resp.text[:200]}"],
                    "raw": resp.text
                }
        except requests.exceptions.ConnectionError:
            return {
                "success": False,
                "created": 0,
                "altered": 0,
                "errors": 1,
                "ignored": 0,
                "messages": [f"Failed to connect to Tally at {self.base_url}. Make sure Tally is running."],
                "raw": ""
            }
        except Exception as e:
            return {
                "success": False,
                "created": 0,
                "altered": 0,
                "errors": 1,
                "ignored": 0,
                "messages": [f"Error communicating with Tally: {str(e)}"],
                "raw": ""
            }

    def _parse_import_response(self, xml_text: str) -> Dict[str, Any]:
        result = {
            "success": False,
            "created": 0,
            "altered": 0,
            "errors": 0,
            "ignored": 0,
            "messages": [],
            "raw": xml_text
        }
        try:
            root = ET.fromstring(xml_text)
            # Find RESPONSE tags
            created = root.find(".//CREATED")
            altered = root.find(".//ALTERED")
            errors = root.find(".//ERRORS")
            ignored = root.find(".//IGNORED")

            if created is not None and created.text:
                result["created"] = int(created.text.strip() or 0)
            if altered is not None and altered.text:
                result["altered"] = int(altered.text.strip() or 0)
            if errors is not None and errors.text:
                result["errors"] = int(errors.text.strip() or 0)
            if ignored is not None and ignored.text:
                result["ignored"] = int(ignored.text.strip() or 0)

            # Collect error messages
            for line_err in root.iter("LINEERROR"):
                if line_err.text and line_err.text.strip():
                    result["messages"].append(line_err.text.strip())

            # Look for other error nodes
            for err_node in root.iter("ERROR"):
                if err_node.text and err_node.text.strip():
                    result["messages"].append(err_node.text.strip())

            for desc in root.iter("DESC"):
                if desc.text and desc.text.strip():
                    result["messages"].append(desc.text.strip())

            # If created or altered > 0 and errors == 0, mark success
            if (result["created"] > 0 or result["altered"] > 0) and result["errors"] == 0:
                result["success"] = True
            elif result["created"] == 0 and result["altered"] == 0 and result["errors"] == 0:
                # Might have ignored or other
                if result["ignored"] > 0:
                    result["messages"].append(f"{result['ignored']} entries were ignored (may already exist).")
                result["success"] = True
            else:
                result["success"] = False

        except ET.ParseError:
            # Check if response contains plain text error
            result["messages"].append(f"Non-XML response received: {xml_text[:300]}")
            result["errors"] = 1
            result["success"] = False

        return result
