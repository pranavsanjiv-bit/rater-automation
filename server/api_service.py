from requests import session
from datetime import datetime as dt
import requests as req
import random as rd
import os
import json
# pyright: ignore [reportMissingImports]
from dotenv import load_dotenv

# Load partner credentials from .env (located next to this file's routes folder)
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "routes", ".env")
load_dotenv(dotenv_path=_env_path)


class APIService:
    def __init__(self):
        self.get_staff_token_url = "https://api-lendpro.sandbox-riskcovry.com/partner_staffs/sso_sign_in.json"
        self.save_loan_det_url = "https://lendproapi-lendpro.sandbox-riskcovry.com/api/partner/save_loan_details"
        self.fetch_loan_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/fetch_loans.json"
        self.save_loan_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/save_loan"
        self.fetch_or_create_lead_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/fetch_or_create_lead"
        self.fetch_customer_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/fetch_customer_details.json?partner_uid={partner_uid}&loan_id={loan_id}"
        self.combo_partner_product_url = "https://api-lendpro.sandbox-riskcovry.com/quotation_searches/combo_partner_products.json"
        self.get_product_quotes_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/v2/get_async_product_quotes.json"
        self.fetch_quote_url = "https://api-lendpro.sandbox-riskcovry.com/loan_protect/v2/fetch_quotes?lead_id={session_token}"
        self.partner_det = {
            "code": os.environ["PARTNER_CODE"],
            "key": os.environ["PARTNER_KEY"],
            "user_id": os.environ["PARTNER_USERID"],
        }


    def gen_app_no(self):
        timestamp = dt.now().strftime("%Y%m%d%H%M%S")
        random_num = rd.randint(100, 999)  # 3-digit random number

        return f"LENDPRO-{timestamp}-{random_num}"
    
    def get_staff_token(self):
        payload = {
            "partner_user_id": self.partner_det["user_id"],
            "quote_id" : ""
        }

        headers = {
            "Partner-Code" : self.partner_det["code"],
        }

        res = req.post(self.get_staff_token_url, json=payload, headers=headers)
        return res.json().get("authentication_token")
    
    def fetch_loan(self, loan_no, partner_code, authToken):
        headers = {
            "Partner-Code": partner_code,
            "staff-token": authToken,
        }

        url = f"{self.fetch_loan_url}?ref_number={loan_no}"

        res = req.get(url=url, headers=headers)
        return res.json()   

    def save_loan(self, save_loan_payload,partner_code, auth_token, loan):
        payload = {
            "loan": {
                "line_of_business": None,
                "branch_code": "HYD001",
                "branch_name": "Hyderabad",
                "tenure": 30,
                "tenure_unit": "years",
                "emi_amount": None,
                "loan_ref_no": None,
                "loan_sanction_amount": None,
                "loan_disbursement_amount": None,
                "loan_applicant_amount": None,
                "loan_sanction_date": None,
                "loan_disbursement_date": None,
                "loan_applicant_date": None,
                "source": None,
                "margin": None,
                "loan_no": loan["loan_account_no"],
                "funding_loan_no": loan["loan_account_no"],
                "type": loan["loan_type"],
                "amount": loan["loan_amount"],
                "interest_rate": "14.00"
            },
            "borrowers": [
            {
                "title": "Mr",
                "first_name": "Vishal",
                "last_name": "Bhat",
                "email": "a@b.com",
                "phone_number": "6361748623",
                "gender": "Transgender",
                "dob": save_loan_payload["borrowers"][0]["dob"],
                "is_primary_borrower": True,
                "user_id": "CUST8018649635",
                "age": 27,
                "annual_income": 1000000,
                "pan": "FPNPK9940H",
                "occupation": "Salaried",
                "marital_status": None,
                "education": None,
                "height": None,
                "weight": None,
                "additional_info": None,
                "address": {
                "address_line_1": "Address Line 1",
                "address_line_2": "Address Line 2",
                "zipcode": "390001",
                "city": "Vadodara",
                "state": "Gujarat",
                "country": None
                },
                "bank_details": {
                "branch_name": None,
                "account_number": None,
                "ifsc": None,
                "account_type": None,
                "bank_name": None,
                "account_holder_name": None,
                "bank_city": None,
                "bank_state": None,
                "bank_district": None,
                "micr_code": None
                },
                "nominees": []
            }
            ],
            "assets": [
            {
                "asset_type": "HOME",
                "owner_name": "XYZ ABC",
                "owner_type": None,
                "owner_pan": None,
                "owner_dob": None,
                "asset_value": None,
                "additional_info": "{\"test\":\"hey\"}",
                "asset_sub_types": [],
                "address": {
                "address_line_1": "J H patel badavane",
                "address_line_2": "Shimoga",
                "zipcode": "577201",
                "city": "Shimoga",
                "state": "Gujarat",
                "country": None
                }
            },
            {
                "asset_type": "BUSINESS",
                "owner_name": "XYZ ABC",
                "owner_type": None,
                "owner_pan": None,
                "owner_dob": None,
                "asset_value": None,
                "additional_info": "{\"test\":\"hey\"}",
                "asset_sub_types": [],
                "address": {
                "address_line_1": "Address Line 1",
                "address_line_2": "Address Line 2",
                "zipcode": "390019",
                "city": "Vadodara",
                "state": "Gujarat",
                "country": None
                }
            }
            ],
            "property": {
            "type": "HOME",
            "address": {
                "address_line_1": "J H patel badavane",
                "address_line_2": "Shimoga",
                "zipcode": "577201",
                "city": "Shimoga",
                "state": "Gujarat"
            }
            }
        }

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token,
            "Entity": "AGENT",
            "staff-phone-no": "7489983860"
        }

        res = req.post(url=self.save_loan_url, headers=headers, json=payload)
        return res.json()
    
    # def save_loan_details(self):
    #     loan_no = self.gen_app_no()

    #     payload = {
    #         "loan": {
    #             "loan_no": loan_no,
    #             "funding_loan_no": loan_no,
    #             "tenure": self.user_det["loan_tenure"],
    #             "tenure_unit": "years",
    #             "type": self.user_det.get("loan_type","HL"),
    #             "amount": self.user_det["loan_amount"],
    #             "interest_rate": 14,
    #             "branch_code": "HYD001",
    #             "branch_name": "Hyderabad", 
    #             "coverage_type" : self.user_det.get("coverage_type", "Reducing")
    #         },
    #         "borrowers": [{
    #             "title": "Mr",
    #             "first_name": 'Vishal',
    #             "last_name": "Bhat",
    #             "phone_number": "6361748623",
    #             "gender": self.user_det.get("gender", "Transgender"),
    #             "dob": self.user_det["dob"],
    #             "email": "a@b.com",
    #             "is_primary_borrower": True,
    #             "partner_uid": "CUST8018649635",
    #             "pan": "FPNPK9940H",
    #             "occupation": "Salaried",
    #             "annual_income": 1000000,
    #             "address": {
    #                 "address_line_1": "Address Line 1",
    #                 "address_line_2": "Address Line 2",
    #                 "city": "Vadodara",
    #                 "state": "Gujarat",
    #                 "zipcode": "390001"
    #             }
    #         }],
    #         "assets": [
    #             {
    #                 "asset_type": "HOME",
    #                 "owner_name": "XYZ ABC",
    #                 "additional_info": {
    #                     "test": "hey"
    #                 },
    #                 "address": {
    #                     "address_line_1": "J H patel badavane",
    #                     "address_line_2": "Shimoga",
    #                     "city": "Shimoga",
    #                     "state": "Gujarat",
    #                     "zipcode": self.user_det.get("property_pincode", "360019")
    #                 },
    #                 "asset_id": "CUST9999999998"
    #             },
    #             {
    #                 "asset_type": "BUSINESS",
    #                 "owner_name": "XYZ ABC",
    #                 "additional_info": {
    #                     "test": "hey"
    #                 },
    #                 "address": {
    #                     "address_line_1": "Address Line 1",
    #                     "address_line_2": "Address Line 2",
    #                     "city": "Vadodara",
    #                     "state": "Gujarat",
    #                     "zipcode": "390019"
    #                 },
    #                 "asset_id": "CUST9999999999"
    #             }
    #         ]
    #     }
    
    #     headers = {
    #         "Partner-Code": self.partner_det["code"],
    #         "Partner-Key": self.partner_det["key"]
    #     }

    #     res = req.post(url=self.save_loan_det_url, headers=headers, json=payload)
    #     res_json = res.json()

    #     if res.status_code == 201 and res_json["loan_no"]:
    #         auth_token = self.get_staff_token()
    #         loan_details = self.fetch_loan(loan_no, self.partner_det["code"], auth_token)
    #         save_loan_response = self.save_loan(payload, self.partner_det["code"], auth_token, loan_details[0])
    #         loan_id = save_loan_response["loan_id"]
    #         fetch_customer_response = self.fetch_customer(payload["borrowers"][0]["partner_uid"], self.partner_det["code"],auth_token, loan_id)
    #         create_lead_response = self.fetch_or_create_lead(payload, self.partner_det["code"], auth_token, payload["loan"], loan_id, fetch_customer_response["borrower_id"]) 
    #         session_token = create_lead_response["session_token"]
    #         combo_partner_product_response = self.combo_partner_product(session_token,  os.getenv("PARTNER_CODE"), auth_token)
    #         get_product_quotes_response = self.get_product_quotes(session_token,  os.getenv("PARTNER_CODE"), auth_token)
    #         fetch_quote_response = self.fetch_quote(session_token, os.getenv("PARTNER_CODE"), auth_token)
    #         quotes = fetch_quote_response.get("quotes", None)
    #         premium = {
    #             "base": quotes[0]["base_premium"] if quotes else -1,
    #             "total" : quotes[0]["total_premium"] if quotes else -1
    #         }

    #         # premium = {
    #         #     "base":-1,
    #         #     "total" :-1
    #         # }
    #         '''
    #             DISPLAY BASE PREMIUM AND TOTAL PREMIUM
    #         '''
    #         # #print(f"{"="*10} API based {'='*10}")
    #         # print(f"Base premium: {premium["base"]}\nTotal premium: {premium["total"]}")
    #         # writeCell("User Inputs", "C8", quotes[0].base_premium)
    #         # writeCell("User Inputs", "C9", quotes[0].total_premium)
    #         result = {
    #             "loan_no":loan_no, 
    #             "premium": premium
    #         }

    #         return result
    
    def fetch_customer(self, partner_uid, partner_code, auth_token, loan_id):
        url = self.fetch_customer_url.format(
            partner_uid = partner_uid,
            loan_id = loan_id
        )

        headers =  {
        "Partner-Code": partner_code,
        "staff-token": auth_token
        }

        res = req.get(url=url, headers=headers)
        return res.json()

    def fetch_or_create_lead(self, save_loan_payload, partner_code, auth_token, loan, loan_id, borrower_id):
        loan_no = loan["loan_no"]
        dob = save_loan_payload["borrowers"][0]["dob"]

        payload = {
            "ref_no": loan_no,
            "line_of_business": None,  
            "loan_account_number": loan_no,
            "tenure": loan["tenure"],
            "tenure_unit": "year",
            "loan_amount": loan["amount"],
            "interest_rate": 14,
            "loan_commencement_date": "",
            "loan_type": "LAP",
            "branch_name": "Hyderabad",
            "branch_code": "HYD001",
            "appilcation_no": loan_no,
            
            "emi_amount": None,
            "loan_ref_no": None,
            "loan_sanction_amount": None,
            "loan_disbursement_amount": None,
            "loan_applicant_amount": None,
            "loan_sanction_date": None,
            "loan_disbursement_date": None,
            "loan_applicant_date": None,
            
            "insurance_loan": {
                "loan_account_number": loan_no
            },
            
            "proposer": {
                "title": "Mr",
                "first_name": "Vishal",
                "last_name": "Bhat",
                "email": "a@b.com",
                "phone_number": "6361748623",
                "gender": "Transgender",
                "dob": dob,
                "pan": "FPNPK9940H",
                "occupation": "Salaried",
                "annual_income": 1000000,
                "is_primary_borrower": True,  
                "partner_uid": "CUST8018649635",
                "address": {
                    "address_line_1": "Address Line 1",
                    "address_line_2": "Address Line 2",
                    "zipcode": "390001",
                    "city": "Vadodara",
                    "state": "Gujarat"
                }
            },
            
            "insured": [
                {
                    "title": "Mr",
                    "first_name": "Vishal",
                    "last_name": "Bhat",
                    "email": "a@b.com",
                    "phone_number": "6361748623",
                    "gender": "Transgender",
                    "dob": dob,
                    "pan": "FPNPK9940H",
                    "occupation": "Salaried",
                    "annual_income": 1000000,
                    "is_primary_borrower": True,
                    "external_user_id": "CUST8018649635",
                    "address": {
                        "address_line_1": "Address Line 1",
                        "address_line_2": "Address Line 2",
                        "zipcode": "390001",
                        "city": "Vadodara",
                        "state": "Gujarat"
                    }
                }
            ],
            
            "property": {
                "type": "HOME",
                "address": {
                    "address_line_1": "J H patel badavane",
                    "address_line_2": "Shimoga",
                    "zipcode": "577201",
                    "city": "Shimoga",
                    "state": "Gujarat"
                }
            },
            
            "loan_id": loan_id,
            "borrower_id": borrower_id
        }   

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token,
            "Entity": "AGENT",
            "staff-phone-no": "7489983860"
        }

        res = req.post(url=self.fetch_or_create_lead_url, headers=headers, json=payload)
        return res.json()
    
    def combo_partner_product(self, session_token, partner_code, auth_token):
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

        res = req.post(url=self.combo_partner_product_url, headers=headers, json=payload)
        return res.json()
    
    def get_product_quotes(self, session_token, partner_code, auth_token, loan_det):
        payload = {
            "lead_id": session_token,
            "plans": [
                {
                    "quote_id": session_token,
                    "product_code": "BAJAJ_LIFE_GCPP",
                    "sum_insured": loan_det['loan_amount'],
                    "tenure": loan_det['tenure']
                }
            ]
        }

        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token
        }

        res = req.post(url=self.get_product_quotes_url, headers=headers, json=payload)
        return res.json()

    def fetch_quote(self, session_token, partner_code, auth_token):
        headers = {
            "Partner-Code": partner_code,
            "staff-token": auth_token
        }

        res = req.get(url=self.fetch_quote_url.format(session_token=session_token), headers=headers)
        return res.json()
        
    def save_loan_details(self,loan_no, payload):
        if "loan" in payload:
            payload["loan"]["loan_no"] = loan_no
            payload["loan"]["funding_loan_no"] = loan_no
    
        headers = {
            "Partner-Code": self.partner_det["code"],
            "Partner-Key": self.partner_det["key"]
        }

        res = req.post(url=self.save_loan_det_url, headers=headers, json=payload)
        res_json = res.json()

        if res.status_code == 201 and res_json["loan_no"]:
            auth_token = self.get_staff_token()
            loan_details = self.fetch_loan(loan_no, self.partner_det["code"], auth_token)
            if not loan_details:
                return {"status": "Failed", "error": "No loan details returned by fetch_loan API."}

            save_loan_response = self.save_loan(payload, self.partner_det["code"], auth_token, loan_details[0])
            if "loan_id" not in save_loan_response:
                return save_loan_response

            loan_id = save_loan_response["loan_id"]
            fetch_customer_response = self.fetch_customer(payload["borrowers"][0]["partner_uid"], self.partner_det["code"], auth_token, loan_id)
            if "borrower_id" not in fetch_customer_response:
                return fetch_customer_response

            create_lead_response = self.fetch_or_create_lead(payload, self.partner_det["code"], auth_token, payload["loan"], loan_id, fetch_customer_response["borrower_id"]) 
            if "session_token" not in create_lead_response:
                return create_lead_response

            session_token = create_lead_response["session_token"]
            combo_partner_product_response = self.combo_partner_product(session_token, os.getenv("PARTNER_CODE"), auth_token)
            get_product_quotes_response = self.get_product_quotes(session_token, os.getenv("PARTNER_CODE"), auth_token, loan_det=loan_details[0])
            fetch_quote_response = self.fetch_quote(session_token, os.getenv("PARTNER_CODE"), auth_token)
            
            quotes = fetch_quote_response.get("quotes")
            if quotes and len(quotes) > 0:
                premium = {
                    "status": "success",
                    "base_premium": quotes[0].get("base_premium"),
                    "total_premium": quotes[0].get("total_premium")
                }
            else:
                premium = {
                    "status": "failed",
                    "base_premium": None,
                    "total_premium": None
                }

            return premium
        else:
            return res_json
    
