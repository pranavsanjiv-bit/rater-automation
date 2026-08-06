from requests import session
from datetime import datetime as dt
import requests as req
import random as rd
import os
import re
import json
import time
# pyright: ignore [reportMissingImports]
from dotenv import load_dotenv

# Load application configuration and partner credentials from the routes/.env file
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "routes", ".env")
load_dotenv(dotenv_path=_env_path)

# ==========================================================
# Initialization
# ==========================================================

class APIService:
    def __init__(self):
        self.get_staff_token_url = os.getenv("GET_STAFF_TOKEN_URL")
        self.save_loan_det_url = os.getenv("SAVE_LOAN_DETAILS_URL")
        self.fetch_loan_url = os.getenv("FETCH_LOAN_URL")
        self.save_loan_url = os.getenv("SAVE_LOAN_URL")
        self.fetch_or_create_lead_url = os.getenv("FETCH_OR_CREATE_LEAD_URL")
        self.fetch_customer_url = os.getenv("FETCH_CUSTOMER_URL")
        self.combo_partner_product_url = os.getenv("COMBO_PARTNER_PRODUCT_URL")
        self.get_product_quotes_url = os.getenv("GET_PRODUCT_QUOTES_URL")
        self.fetch_quote_url = os.getenv("FETCH_QUOTE_URL")
        self.partner_det = {
            "code": os.environ["PARTNER_CODE"],
            "key": os.environ["PARTNER_KEY"],
            "user_id": os.environ["PARTNER_USERID"],
        }
        self.quote_poll_interval = float(os.getenv("PARTNER_QUOTE_POLL_INTERVAL", "1.5"))

# =====================================================
# Authentication APIs
# =====================================================
    def gen_app_no(self):
        timestamp = dt.now().strftime("%Y%m%d%H%M%S")
        random_num = rd.randint(100, 999)  # 3-digit random number

        return f"LENDPRO-{timestamp}-{random_num}"

    def get_staff_token(self, timeout=None):
        """Generate and return the staff authentication token."""
        payload = {
            "partner_user_id": self.partner_det["user_id"],
            "quote_id" : ""
        }

        headers = {
            "Partner-Code" : self.partner_det["code"],
        }

        res = self._safe_post(self.get_staff_token_url, json=payload, headers=headers, timeout=timeout)
        return res.json().get("authentication_token")

    def fetch_loan(self, loan_no, partner_code, authToken, timeout=None):
        """Fetch loan details using the generated loan number."""
        headers = {
            "Partner-Code": partner_code,
            "staff-token": authToken,
        }

        url = f"{self.fetch_loan_url}?ref_number={loan_no}"

        res = self._safe_get(url=url, headers=headers, timeout=timeout)
        return res.json()

    def _safe_post(self, url, headers=None, json=None, timeout=None):
        try:
            return req.post(url=url, headers=headers, json=json, timeout=timeout)
        except req.exceptions.Timeout as exc:
            raise TimeoutError(f"Partner API request timed out for {url}") from exc
        except req.exceptions.RequestException as exc:
            raise ConnectionError(f"Partner API request failed for {url}: {exc}") from exc

    def _safe_get(self, url, headers=None, timeout=None):
        try:
            return req.get(url=url, headers=headers, timeout=timeout)
        except req.exceptions.Timeout as exc:
            raise TimeoutError(f"Partner API request timed out for {url}") from exc
        except req.exceptions.RequestException as exc:
            raise ConnectionError(f"Partner API request failed for {url}: {exc}") from exc

# ---------- Helper Methods ----------
    def _normalize_numeric_value(self, value):
        if value is None:
            return None
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value) if value.is_integer() else value
        if isinstance(value, str):
            try:
                cleaned = value.strip()
                if cleaned == "":
                    return None
                # Accept comma separators and common currency prefixes/suffixes.
                cleaned = re.sub(r"[₹Rs\.\s,]+", "", cleaned, flags=re.IGNORECASE)
                if cleaned == "":
                    return None
                if "." in cleaned:
                    parsed = float(cleaned)
                    return int(parsed) if parsed.is_integer() else parsed
                return int(cleaned)
            except ValueError:
                return value
        return value

    def _normalize_tenure_to_years(self, tenure, tenure_unit):
        tenure = self._normalize_numeric_value(tenure)
        if tenure is None:
            return None

        if isinstance(tenure_unit, str):
            normalized_unit = tenure_unit.strip().lower()
        else:
            normalized_unit = "years"

        if normalized_unit in {"months", "month", "m"}:
            years = float(tenure) / 12.0
            return int(years) if years.is_integer() else years

        return tenure

    def _normalize_boolean(self, value, default=None):
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            lowered = value.strip().lower()
            if lowered in {"true", "1", "yes", "y"}:
                return True
            if lowered in {"false", "0", "no", "n"}:
                return False
        return default if value is None else value

    def _normalize_dob(self, dob_value):
        if isinstance(dob_value, str):
            dob_value = dob_value.strip()
            if "/" in dob_value:
                dob_value = dob_value.replace("/", "-")
        return dob_value

    def _normalize_text(self, value, max_len=100, default="NA"):
        if value is None:
            return default
        if not isinstance(value, str):
            value = str(value)
        cleaned = "".join(ch for ch in value.strip() if ord(ch) >= 32)
        if not cleaned:
            return default
        if len(cleaned) > max_len:
            cleaned = cleaned[:max_len]
        return cleaned

    def _normalize_address(self, address: dict) -> dict:
        if not isinstance(address, dict):
            return {
                "address_line_1": "Address Line 1",
                "address_line_2": "Address Line 2",
                "city": "Unknown",
                "state": "Unknown",
                "zipcode": "000000",
            }
        address = address.copy()
        address["address_line_1"] = self._normalize_text(address.get("address_line_1"), max_len=100, default="Address Line 1")
        address["address_line_2"] = self._normalize_text(address.get("address_line_2"), max_len=100, default="Address Line 2")
        address["city"] = self._normalize_text(address.get("city"), max_len=50, default="Unknown")
        address["state"] = self._normalize_text(address.get("state"), max_len=50, default="Unknown")
        address["zipcode"] = self._normalize_text(address.get("zipcode"), max_len=20, default="000000")
        return address

    def _normalize_occupation(self, value):
        if value is None:
            return "Others"
        if not isinstance(value, str):
            value = str(value)
        normalized = value.strip().title()
        mapping = {
            "Self Employed": "Self Employed",
            "Self-Employed": "Self Employed",
            "Business": "Self Employed",
            "Salaried": "Salaried",
            "Student": "Student",
            "Housewife": "Housewife",
            "Other": "Others",
            "Others": "Others",
            "Self": "Self Employed",
        }
        return mapping.get(normalized, "Others")

    def _normalize_borrower(self, borrower: dict, index: int = 0) -> dict:
        if not isinstance(borrower, dict):
            return borrower

        borrower = borrower.copy()
        borrower["age"] = self._normalize_numeric_value(borrower.get("age"))
        borrower["annual_income"] = self._normalize_numeric_value(borrower.get("annual_income"))
        borrower["height"] = self._normalize_numeric_value(borrower.get("height"))
        borrower["weight"] = self._normalize_numeric_value(borrower.get("weight"))
        borrower["is_primary_borrower"] = self._normalize_boolean(
            borrower.get("is_primary_borrower"), default=(index == 0)
        )
        borrower["dob"] = self._normalize_dob(borrower.get("dob"))
        borrower["occupation"] = self._normalize_occupation(borrower.get("occupation"))

        if "partner_uid" not in borrower or not borrower["partner_uid"]:
            borrower["partner_uid"] = borrower.get("user_id") or f"CUST{int(dt.now().timestamp())}"
        if "user_id" not in borrower or not borrower["user_id"]:
            borrower["user_id"] = borrower.get("partner_uid")

        address = borrower.get("address") or {}
        borrower["address"] = self._normalize_address(address)

        return borrower

    def _normalize_lead_payload(self, payload: dict) -> dict:
        candidate = payload.get("loan_account_number") or payload.get("loan_number") or payload.get("application_no") or payload.get("funding_loan_number")
        loan = payload.get("loan", {})
        candidate = candidate or loan.get("application_no") or loan.get("loan_number") or loan.get("loan_no") or loan.get("funding_loan_no")

        if candidate:
            payload["loan_account_number"] = payload.get("loan_account_number") or candidate
            payload["loan_number"] = payload.get("loan_number") or candidate
            payload["funding_loan_number"] = payload.get("funding_loan_number") or candidate

        if "ref_no" not in payload or not payload.get("ref_no"):
            payload["ref_no"] = payload.get("loan_account_number") or payload.get("loan_number") or payload.get("application_no") or loan.get("application_no")

        if "insurance_loan" not in payload or not payload.get("insurance_loan"):
            loan_acc = payload.get("loan_account_number")
            if loan_acc:
                payload["insurance_loan"] = {"loan_account_number": loan_acc}

        if "appilcation_no" not in payload and payload.get("loan", {}).get("application_no"):
            payload["appilcation_no"] = payload["loan"]["application_no"]

        if "insured" not in payload or not payload["insured"]:
            borrowers = payload.get("borrowers") or []
            if borrowers:
                insured_person = borrowers[0].copy()
                insured_person["external_user_id"] = insured_person.get("partner_uid") or insured_person.get("user_id")
                if "address" in insured_person:
                    insured_person["address"] = self._normalize_address(insured_person["address"])
                payload["insured"] = [insured_person]

        if "loan" not in payload:
            loan_number = payload.pop("loan_number", None)
            funding_loan_number = payload.pop("funding_loan_number", None)
            payload["loan"] = {
                "application_no": loan_number,
                "funding_loan_no": funding_loan_number,
                "amount": self._normalize_numeric_value(payload.pop("loan_amount", None)),
                "tenure": self._normalize_numeric_value(payload.pop("loan_tenure", None)),
                "tenure_unit": payload.pop("loan_tenure_unit", None),
                "type": payload.pop("loan_type", None),
                "interest_rate": self._normalize_numeric_value(payload.pop("interest_rate", None)),
                "branch_code": payload.pop("branch_code", None),
                "branch_name": payload.pop("branch_name", "Hyderabad"),
            }
        else:
            payload["loan"]["amount"] = self._normalize_numeric_value(payload["loan"].get("amount"))
            payload["loan"]["tenure"] = self._normalize_numeric_value(payload["loan"].get("tenure"))
            payload["loan"]["interest_rate"] = self._normalize_numeric_value(payload["loan"].get("interest_rate"))

        if "loan_account_number" not in payload:
            payload["loan_account_number"] = payload["loan"].get("application_no") or payload["loan"].get("funding_loan_no")

        borrowers = payload.get("borrowers") or []
        normalized_borrowers = []
        for idx, borrower in enumerate(borrowers):
            normalized_borrowers.append(self._normalize_borrower(borrower, index=idx))
        if normalized_borrowers:
            payload["borrowers"] = normalized_borrowers

        if "proposer" in payload and isinstance(payload["proposer"], dict):
            payload["proposer"]["address"] = self._normalize_address(payload["proposer"].get("address") or {})

        if "insured" in payload and isinstance(payload["insured"], list) and payload["insured"]:
            insured_item = payload["insured"][0]
            if isinstance(insured_item, dict) and "address" in insured_item:
                insured_item["address"] = self._normalize_address(insured_item["address"] or {})
                payload["insured"][0] = insured_item

        if "borrowers" in payload and payload["borrowers"]:
            borrower = payload["borrowers"][0]
            if "partner_uid" not in borrower or not borrower["partner_uid"]:
                borrower["partner_uid"] = borrower.get("partner_uid") or borrower.get("user_id") or f"CUST{int(dt.now().timestamp())}"

        return payload

    def _build_save_loan_borrower(self, borrower: dict) -> dict:
        partner_uid = borrower.get("partner_uid") or borrower.get("user_id") or f"CUST{int(dt.now().timestamp())}"
        user_id = borrower.get("user_id") or borrower.get("partner_uid") or partner_uid

        return {
            "title": borrower.get("title"),
            "first_name": borrower.get("first_name"),
            "last_name": borrower.get("last_name"),
            "email": borrower.get("email"),
            "phone_number": borrower.get("phone_number"),
            "gender": borrower.get("gender"),
            "dob": borrower.get("dob"),
            "is_primary_borrower": borrower.get("is_primary_borrower", True),
            "user_id": user_id,
            "partner_uid": partner_uid,
            "age": self._normalize_numeric_value(borrower.get("age")),
            "annual_income": self._normalize_numeric_value(borrower.get("annual_income")),
            "pan": borrower.get("pan"),
            "occupation": borrower.get("occupation"),
            "marital_status": borrower.get("marital_status"),
            "education": borrower.get("education"),
            "height": self._normalize_numeric_value(borrower.get("height")),
            "weight": self._normalize_numeric_value(borrower.get("weight")),
            "additional_info": borrower.get("additional_info"),
            "borrower_type": borrower.get("borrower_type", "Individual"),
            "address": borrower.get("address", {}),
            "bank_details": borrower.get("bank_details", {}),
            "nominees": borrower.get("nominees", []),
        }

    def _normalize_tenure_unit(self, tenure_unit):
        if isinstance(tenure_unit, str):
            normalized_unit = tenure_unit.strip().lower()
        else:
            normalized_unit = "years"

        if normalized_unit in {"months", "month", "m"}:
            return "months"
        return "years"

    def _convert_tenure(self, tenure, from_unit, to_unit):
        if tenure is None:
            return None
        if from_unit == to_unit:
            return tenure
        if from_unit == "months" and to_unit == "years":
            years = float(tenure) / 12.0
            return int(years) if years.is_integer() else years
        if from_unit == "years" and to_unit == "months":
            return int(float(tenure) * 12)
        return tenure

    def _build_quote_payload(self, session_token, sum_insured, tenure, tenure_unit):
        return {
            "lead_id": session_token,
            "plans": [
                {
                    "quote_id": session_token,
                    "product_code": "BAJAJ_LIFE_GCPP",
                    "sum_insured": sum_insured,
                    "tenure": tenure,
                    "tenure_unit": tenure_unit,
                    "loan_tenure": tenure,
                    "loan_tenure_unit": tenure_unit,
                    "policy_tenure": tenure,
                    "policy_tenure_unit": tenure_unit,
                }
            ]
        }

    def _policy_tenure_validation_failed(self, response_data):
        if not isinstance(response_data, dict):
            return False
        quotes = response_data.get("quotes") or []
        for quote in quotes:
            if quote.get("status") == "Failed":
                message = quote.get("error_message") or quote.get("message") or ""
                if isinstance(message, str) and "Policy Tenure should be equal to loan tenure" in message:
                    return True
        return False


# ---------- Business Methods ----------
# =====================================================
# Loan APIs
# =====================================================
    def save_loan(self, save_loan_payload, partner_code, auth_token, loan, timeout=None):
        """Save loan information in Loan Protect."""
        borrower_payload = (save_loan_payload.get("borrowers") or [{}])[0]
        borrower = self._build_save_loan_borrower(borrower_payload)

        loan_account_no = loan.get("loan_account_no") or loan.get("loan_no") or loan.get("application_no")
        payload = {
            "loan": {
                "line_of_business": None,
                "branch_code": loan.get("branch_code", "HYD001"),
                "branch_name": loan.get("branch_name", "Hyderabad"),
                "tenure": loan.get("tenure"),
                "tenure_unit": loan.get("tenure_unit", "years"),
                "emi_amount": loan.get("emi_amount"),
                "loan_ref_no": loan.get("loan_ref_no"),
                "loan_sanction_amount": loan.get("loan_sanction_amount"),
                "loan_disbursement_amount": loan.get("loan_disbursement_amount"),
                "loan_applicant_amount": loan.get("loan_applicant_amount"),
                "loan_sanction_date": loan.get("loan_sanction_date"),
                "loan_disbursement_date": loan.get("loan_disbursement_date"),
                "loan_applicant_date": loan.get("loan_applicant_date"),
                "source": loan.get("source"),
                "margin": loan.get("margin"),
                "loan_no": loan_account_no,
                "funding_loan_no": loan_account_no,
                "type": loan.get("loan_type") or loan.get("type"),
                "amount": loan.get("loan_amount") or loan.get("amount"),
                "interest_rate": str(loan.get("interest_rate", "14.00")),
            },
            "borrowers": [borrower],
            "assets": save_loan_payload.get("assets", []),
            "property": save_loan_payload.get("property", {}),
        }

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token,
            "Entity": "AGENT",
            "staff-phone-no": "7489983860"
        }

        res = self._safe_post(url=self.save_loan_url, headers=headers, json=payload, timeout=timeout)
        return res.json()

# =====================================================
# Lead APIs
# =====================================================
    def fetch_customer(self, partner_uid, partner_code, auth_token, loan_id, timeout=None):
        """Fetch customer details for the specified loan."""
        url = self.fetch_customer_url.format(
            partner_uid = partner_uid,
            loan_id = loan_id
        )

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token
        }

        res = self._safe_get(url=url, headers=headers, timeout=timeout)
        return res.json()

    def fetch_or_create_lead(self, save_loan_payload, partner_code, auth_token, loan, loan_id, borrower_id, timeout=None):
        """Fetch an existing lead or create a new one."""
        borrower = (save_loan_payload.get("borrowers") or [{}])[0]
        partner_uid = borrower.get("partner_uid") or borrower.get("user_id") or f"CUST{int(dt.now().timestamp())}"
        user_id = borrower.get("user_id") or borrower.get("partner_uid") or partner_uid
        dob = borrower.get("dob")

        loan_account_no = loan.get("loan_account_no") or loan.get("loan_no") or loan.get("application_no")
        loan_type = loan.get("loan_type") or loan.get("type") or save_loan_payload.get("loan", {}).get("type") or save_loan_payload.get("loan_type")
        loan_amount = loan.get("loan_amount") or loan.get("amount")

        proposer = {
            "title": borrower.get("title"),
            "first_name": borrower.get("first_name"),
            "last_name": borrower.get("last_name"),
            "email": borrower.get("email"),
            "phone_number": borrower.get("phone_number"),
            "gender": borrower.get("gender"),
            "dob": dob,
            "pan": borrower.get("pan"),
            "occupation": borrower.get("occupation"),
            "annual_income": borrower.get("annual_income"),
            "is_primary_borrower": borrower.get("is_primary_borrower", True),
            "partner_uid": partner_uid,
            "user_id": user_id,
            "address": borrower.get("address", {}),
        }

        insured_person = proposer.copy()
        insured_person["external_user_id"] = user_id

        property_payload = save_loan_payload.get("property")
        if not property_payload:
            property_payload = {
                "type": "HOME",
                "address": borrower.get("address", {}),
            }

        tenure = loan.get("tenure")
        tenure_unit = "years"

        payload = {
            "ref_no": loan_account_no,
            "line_of_business": None,
            "loan_account_number": loan_account_no,
            "tenure": tenure,
            "tenure_unit": tenure_unit,
            "loan_tenure": tenure,
            "loan_tenure_unit": tenure_unit,
            "policy_tenure": tenure,
            "policy_tenure_unit": tenure_unit,
            "loan_amount": loan_amount,
            "interest_rate": loan.get("interest_rate") or 14,
            "loan_commencement_date": save_loan_payload.get("loan_commencement_date", ""),
            "loan_type": loan_type,
            "branch_name": loan.get("branch_name") or save_loan_payload.get("branch_name", "Hyderabad"),
            "branch_code": loan.get("branch_code") or save_loan_payload.get("branch_code", "HYD001"),
            "appilcation_no": loan_account_no,
            "emi_amount": loan.get("emi_amount"),
            "loan_ref_no": loan.get("loan_ref_no"),
            "loan_sanction_amount": loan.get("loan_sanction_amount"),
            "loan_disbursement_amount": loan.get("loan_disbursement_amount"),
            "loan_applicant_amount": loan.get("loan_applicant_amount"),
            "loan_sanction_date": loan.get("loan_sanction_date"),
            "loan_disbursement_date": loan.get("loan_disbursement_date"),
            "loan_applicant_date": loan.get("loan_applicant_date"),
            "insurance_loan": {
                "loan_account_number": loan_account_no
            },
            "proposer": proposer,
            "insured": [insured_person],
            "property": property_payload,
            "loan_id": loan_id,
            "borrower_id": borrower_id,
        }

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token,
            "Entity": "AGENT",
            "staff-phone-no": "7489983860"
        }

        res = self._safe_post(url=self.fetch_or_create_lead_url, headers=headers, json=payload, timeout=timeout)
        return res.json()

    def create_lead(self, lead_payload, partner_code):
        """Create a new lead and initiate the quotation workflow."""
        payload = self._normalize_lead_payload(lead_payload)
        loan_no = self.gen_app_no()
        return self.save_loan_details(loan_no, payload)

    def save_lead(self, lead_payload, partner_code):
        """Save a lead and initiate the quotation workflow."""
        payload = self._normalize_lead_payload(lead_payload)
        loan_no = self.gen_app_no()
        return self.save_loan_details(loan_no, payload)

# =====================================================
# Quote APIs
# =====================================================
    def combo_partner_product(self, session_token, partner_code, auth_token, timeout=None):
        """Fetch partner-product combinations for the generated lead."""
        payload = {
            "session_token": session_token,
            "selected_products": [
                {
                "product_code": "BAJAJ_LIFE_GCPP",
                "isMainSelectedProductInRecommended": False
                }
            ]
        }
        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token,
            "entity": "AGENT",
            "staff-phone-no": "7489983860"
        }

        res = self._safe_post(url=self.combo_partner_product_url, headers=headers, json=payload, timeout=timeout)
        return res.json()

    def get_product_quotes(self, session_token, partner_code, auth_token, loan_det, timeout=None):
        """Request product quotes for the generated lead."""
        tenure = self._normalize_numeric_value(loan_det.get('tenure'))
        tenure_unit = "years"

        sum_insured = self._normalize_numeric_value(
            loan_det.get('loan_amount')
            or loan_det.get('amount')
            or loan_det.get('loan_sanction_amount')
            or loan_det.get('loan_applicant_amount')
        )

        payload = self._build_quote_payload(session_token, sum_insured, tenure, tenure_unit)

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token
        }

        res = self._safe_post(url=self.get_product_quotes_url, headers=headers, json=payload, timeout=timeout)
        response_data = res.json()

        response_data["quote_request_payload"] = payload
        return response_data

    def fetch_quote(self, session_token, partner_code, auth_token, timeout=None):
        """Fetch the final premium quote for the generated lead."""
        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token
        }

        res = self._safe_get(url=self.fetch_quote_url.format(session_token=session_token), headers=headers, timeout=timeout)
        return res.json()

# ==========================================================
# Workflow
# ==========================================================
    def _extract_partner_premium(self, quote: dict):
        if quote is None:
            return None, None

        base_keys = ["base_premium", "basePremium", "premium", "amount", "quote_amount", "net_premium", "premium_amount", "baseAmount"]
        total_keys = ["total_premium", "totalPremium", "total_amount", "gross_premium", "premium_incl_tax", "total_amount_with_gst", "total_amount_incl_tax", "total_amount_with_gst", "amount_with_gst"]

        def _find_nested_value(source, keys):
            if isinstance(source, dict):
                for key, val in source.items():
                    if key in keys and val not in (None, ""):
                        return self._normalize_numeric_value(val)
                    nested = _find_nested_value(val, keys)
                    if nested is not None:
                        return nested
            elif isinstance(source, list):
                for item in source:
                    nested = _find_nested_value(item, keys)
                    if nested is not None:
                        return nested
            return None

        def _find_candidate(source, contains):
            if isinstance(source, dict):
                for key, val in source.items():
                    if contains in key and val not in (None, "") and isinstance(val, (int, float, str)):
                        normalized = self._normalize_numeric_value(val)
                        if isinstance(normalized, (int, float)):
                            return normalized
                    nested = _find_candidate(val, contains)
                    if nested is not None:
                        return nested
            elif isinstance(source, list):
                for item in source:
                    nested = _find_candidate(item, contains)
                    if nested is not None:
                        return nested
            return None

        base = _find_nested_value(quote, base_keys)
        total = _find_nested_value(quote, total_keys)

        if base is None:
            base = _find_candidate(quote, "premium") or _find_candidate(quote, "amount")
        if total is None:
            total = _find_candidate(quote, "total") or _find_candidate(quote, "gross") or _find_candidate(quote, "incl")

        if total is None and base is not None and isinstance(base, (int, float)):
            total = round(base * 1.18, 2)
        return base, total

    def _extract_quote_status(self, response: dict | None) -> str | None:
        if not isinstance(response, dict):
            return None

        quotes = response.get("quotes")
        if isinstance(quotes, dict):
            quotes = [quotes]
        if isinstance(quotes, list) and quotes:
            first_quote = quotes[0]
            if isinstance(first_quote, dict):
                status = first_quote.get("status")
                if isinstance(status, str) and status.strip():
                    return status.strip().lower()

        status = response.get("status")
        if isinstance(status, str) and status.strip():
            return status.strip().lower()

        return None

    def _is_quote_response_pending(self, response: dict | None) -> bool:
        status = self._extract_quote_status(response)
        if status in {"pending", "processing", "in_progress", "queued", "waiting", "requested", "initiated", "received", "inprogress"}:
            return True

        if isinstance(response, dict) and response.get("quotes") is not None:
            quotes = response.get("quotes")
            if isinstance(quotes, list) and len(quotes) == 0:
                return True

        return False

    def _poll_fetch_quote(self, session_token, partner_code, auth_token, interval=None):
        if interval is None:
            interval = self.quote_poll_interval

        while True:
            try:
                response = self.fetch_quote(
                    session_token,
                    partner_code,
                    auth_token,
                    timeout=None,
                )
            except (TimeoutError, ConnectionError) as exc:
                time.sleep(max(0.1, interval))
                continue

            if not self._is_quote_response_pending(response):
                return response

            time.sleep(max(0.1, interval))

    def save_loan_details(self,loan_no, payload):
        """Execute the end-to-end loan processing workflow."""
        if "loan" in payload:
            payload["loan"]["loan_no"] = loan_no
            payload["loan"]["funding_loan_no"] = loan_no

        headers = {
            "Partner-Code": self.partner_det["code"],
            "Partner-Key": self.partner_det["key"]
        }

        try:
            res = self._safe_post(url=self.save_loan_det_url, headers=headers, json=payload,
                                   timeout=None)
            res_json = res.json()
            raw_save_loan_response = res_json
            raw_save_loan_status_code = res.status_code

            if res.status_code == 201 and res_json.get("loan_no"):
                auth_token = self.get_staff_token(timeout=None)
                loan_details = self.fetch_loan(loan_no, self.partner_det["code"], auth_token,
                                               timeout=None)
                if not loan_details:
                    return {"status": "failed", "error": "No loan details returned by fetch_loan API."}

                save_loan_response = self.save_loan(payload, self.partner_det["code"], auth_token, loan_details[0],
                                                   timeout=None)
                if "loan_id" not in save_loan_response:
                    return save_loan_response

                loan_id = save_loan_response["loan_id"]
                fetch_customer_response = self.fetch_customer(payload["borrowers"][0]["partner_uid"],
                                                             self.partner_det["code"], auth_token, loan_id,
                                                             timeout=None)
                if "borrower_id" not in fetch_customer_response:
                    return fetch_customer_response

                create_lead_response = self.fetch_or_create_lead(payload, self.partner_det["code"], auth_token,
                                                                payload["loan"], loan_id,
                                                                fetch_customer_response["borrower_id"],
                                                                timeout=None)
                session_token = create_lead_response.get("session_token") or create_lead_response.get("lead_id") or create_lead_response.get("transaction_id") or create_lead_response.get("loan_id")
                if not session_token:
                    return create_lead_response

                create_lead_response["session_token"] = session_token
                combo_partner_product_response = self.combo_partner_product(session_token, os.getenv("PARTNER_CODE"), auth_token,
                                                                            timeout=None)
                get_product_quotes_response = self.get_product_quotes(session_token, os.getenv("PARTNER_CODE"), auth_token,
                                                                      loan_det=loan_details[0],
                                                                      timeout=None)
                fetch_quote_response = self.fetch_quote(session_token, os.getenv("PARTNER_CODE"), auth_token,
                                                         timeout=None)

                if self._is_quote_response_pending(fetch_quote_response):
                    fetch_quote_response = self._poll_fetch_quote(
                        session_token,
                        os.getenv("PARTNER_CODE"),
                        auth_token,
                    )

                quotes = fetch_quote_response.get("quotes") if isinstance(fetch_quote_response, dict) else None
                if isinstance(quotes, dict):
                    quotes = [quotes]
                if quotes is None:
                    quotes = []
                if not isinstance(quotes, list):
                    quotes = [quotes]

                first_quote = quotes[0] if quotes else None
                quote_status = None
                if isinstance(first_quote, dict):
                    quote_status = str(first_quote.get("status", "")).strip().lower() or None

                base_premium, total_premium = self._extract_partner_premium(first_quote) if first_quote else (None, None)
                if base_premium is None and total_premium is None:
                    base_premium, total_premium = self._extract_partner_premium(fetch_quote_response)
                if base_premium is None and total_premium is None:
                    base_premium, total_premium = self._extract_partner_premium(get_product_quotes_response)

                quote_errors = None
                if isinstance(fetch_quote_response, dict):
                    quote_errors = fetch_quote_response.get("errors") or fetch_quote_response.get("error")
                if not quote_errors and isinstance(get_product_quotes_response, dict):
                    quote_errors = get_product_quotes_response.get("errors") or get_product_quotes_response.get("error")

                premium = {
                    "session_token": session_token,
                    "loan_id": loan_id,
                    "raw_get_product_quotes": get_product_quotes_response,
                    "raw_fetch_quote": fetch_quote_response,
                }

                if quote_status != "failed" and (base_premium is not None or total_premium is not None):
                    premium.update({
                        "status": "success",
                        "base_premium": base_premium,
                        "total_premium": total_premium,
                    })
                else:
                    premium.update({
                        "status": "failed",
                        "base_premium": base_premium,
                        "total_premium": total_premium,
                        "errors": quote_errors,
                    })

                return premium
        except TimeoutError as exc:
            return {
                "status": "failed",
                "error": str(exc),
            }
        except ConnectionError as exc:
            return {
                "status": "failed",
                "error": str(exc),
            }
        else:
            return {
                "status": "failed",
                "error": (
                    res_json.get("error")
                    or res_json.get("message")
                    or f"Initial save loan request failed with status {res.status_code}."
                ),
                "raw_save_loan_response": raw_save_loan_response,
                "raw_save_loan_status_code": raw_save_loan_status_code,
            }

