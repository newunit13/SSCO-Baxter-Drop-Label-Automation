
from typing import Dict, List
from datetime import datetime
import extract_msg
import csv
import sys
import re

def ParseShipperAddress(shipper: str) -> Dict:

    data = [elem.strip() for elem in shipper.split('\n')]

    shipper = {
        "recipent_line_1"   : data[0],
        "recipent_line_2"   : "" if data[1][0].isnumeric() else data[1],
        "address_line_1"    : data[len(data)-5] if data[-5][0].isnumeric() else data[-6],
        "address_line_2"    : '' if data[-5][0].isnumeric() else data[-5],
        "city"              : data[-4].split(", ")[0].strip(),
        "state"             : data[-4].split(", ")[1].split('  ')[0],
        "zip"               : data[-4].split(", ")[1].split('  ')[1],
        "country"           : data[-4].split(", ")[1].split('  ')[2],
        "contact"           : data[-3].split(": ")[1],
        "phone"             : data[-2],
        "email"             : data[-1]
    }

    return shipper

def ParseConsigneeAddress(consignee: str) -> Dict:

    data = [elem.strip() for elem in consignee.split('\n')]

    if data[-2].startswith("Contact"):
        consignee = {
            "recipent_line_1"   : data[0],
            "recipent_line_2"   : "" if data[1][0].isnumeric() else data[1],
            "address_line_1"    : data[-4] if data[-4][0].isnumeric() else data[-5],
            "address_line_2"    : '' if data[-4][0].isnumeric() else data[-4],
            "city"              : data[-3].split(", ")[0].strip(),
            "state"             : data[-3].split(", ")[1].split('  ')[0],
            "zip"               : data[-3].split(", ")[1].split('  ')[1],
            "country"           : data[-3].split(", ")[1].split('  ')[2],
            "phone"             : data[-1],
        }
    else:
        consignee = {
            "recipent_line_1"   : data[0],
            "recipent_line_2"   : "" if data[1][0].isnumeric() else data[1],
            "address_line_1"    : data[-3] if data[-3][0].isnumeric() else data[-4],
            "address_line_2"    : '' if data[-3][0].isnumeric() else data[-3],
            "city"              : data[-2].split(", ")[0].strip(),
            "state"             : data[-2].split(", ")[1].split('  ')[0],
            "zip"               : data[-2].split(", ")[1].split('  ')[1],
            "country"           : data[-2].split(", ")[1].split('  ')[2],
            "phone"             : data[-1],
        }

        if 'null' in consignee["country"]:
            consignee["country"] = 'US'

    return consignee

def ParsePickupLocationAddress(pickup_location: str) -> Dict:

    data = [elem.strip() for elem in pickup_location.split('\n')[:-3]]

    pickup_location = {
        "recipent_line_1"   : data[0],
        "recipent_line_2"   : "" if data[1][0].isnumeric() else data[1],
        "address_line_1"    : data[-5] if data[-5][0].isnumeric() else data[-6],
        "address_line_2"    : '' if data[-5][0].isnumeric() else data[-5],
        "city"              : data[-4].split(", ")[0].strip(),
        "state"             : data[-4].split(", ")[1].split('  ')[0],
        "zip"               : data[-4].split(", ")[1].split('  ')[1],
        "country"           : data[-4].split(", ")[1].split('  ')[2],
        "contact"           : data[-3].split(": ")[1],
        "phone"             : data[-2],
        "email"             : data[-1]
    }

    return pickup_location

def ProcessDrop(input_file: str) -> Dict:

    results = {
        'sender'    : '',
        'to'        : [],
        'subject'   : '',
        'sendDate'  : '',
        'numDrops'  : 0,
        'amiaDrops' : 0,
        'baxDrops'  : 0,
        'successess': {},
        'failures'  : {}
    }

    # DEBUG
    if input_file.endswith(".tsv"):
        tsv = open(f'{input_file}', 'r', newline='')
        results["successess"] = {record["ReferenceNumber"]: record for record in csv.DictReader(tsv, delimiter='\t')}
        results["sender"] = 'debug'
        results["to"] = 'debug'
        results["subject"] = 'debug'
        results['sendDate'] = '1/1/1900'
        results["numDrops"] = len(results["successess"])
        results["amiaDrops"] = len([x for x in results["successess"].values() if x.get("SKU") == "DROP - AMIA"])
        results["baxDrops"] = len(results["successess"]) - results["amiaDrops"]

        for rn, record in enumerate(results["successess"].values()):
            record["shipper"] = {"debug": f"debug {rn}"}
            record["shipperRaw"] = f"debug {rn}"
            record["consignee"] = {"debug": f"debug {rn}"}
            record["consigneeRaw"] = f"debug {rn}"
            record["pickupLocation"] = {"debug": f"debug {rn}"}
            record["pickupLocationRaw"] = f"debug {rn}"
            
        tsv.close()
        return results

    
    msg = extract_msg.Message(input_file)

    results["sender"]   = msg.sender
    results["to"]       = [recp.name for recp in msg.recipients]
    results["subject"]  = msg.subject
    results["sendDate"] = msg.date
    results["numDrops"] = len(msg.attachments)


    if results["numDrops"] > 0:
        print(f"Processing {results['numDrops']} bookings")
        for i, fl in enumerate(msg.attachments):

            booking_id = re.findall(r"Booking ID: (.*)", fl.data.body)[0].strip().replace('\r','')
            shipper = ParseShipperAddress(re.findall(r"SHIPPER\s([\s\S]*)\s{2}CONSIGNEE", fl.data.body)[0].strip().replace('\r',''))
            consignee = ParseConsigneeAddress(re.findall(r"CONSIGNEE\s([\s\S]*)\s{2}PICKUP LOCATION", fl.data.body)[0].strip().replace('\r',''))
            pickup_location = ParsePickupLocationAddress(re.findall(r"PICKUP LOCATION\s([\s\S]*)INSTRUCTIONS", fl.data.body)[0].strip().replace('\r',''))

            try:

                shipment_info = {
                    "ReferenceNumber"                   : booking_id,
                    "PurchaseOrderNumber"               : "",
                    "ShipCarrier"                       : "UPS",
                    "ShipService"                       : re.findall(r"SERVICE REQUESTED:\s+(.*)", fl.data.body)[0].replace('\r','').replace('\n',''),
                    "ShipBilling"                       : re.findall(r"Origin Charges:\s+(.*)", fl.data.body)[0].replace('\r','').replace('\n',''),
                    "ShipAccount"                       : "",
                    "ShipDate"                          : "",
                    "CancelDate"                        : "",
                    "Notes"                             : "",
                    "ShipTo Name"                       : consignee["recipent_line_1"],
                    "ShipToCompany"                     : consignee["recipent_line_2"],
                    "ShipToAddress1"                    : consignee["address_line_1"],
                    "ShipToAddress2"                    : consignee["address_line_2"],
                    "ShipToCity"                        : consignee["city"],
                    "ShipToState"                       : consignee["state"],
                    "ShipToZip"                         : consignee["zip"],
                    "ShipToCountry"                     : consignee["country"],
                    "ShipToPhone"                       : consignee["phone"],
                    "ShipToFax"                         : "",
                    "ShipToEmail"                       : "",
                    "ShipToCustomerID"                  : "",
                    "ShipToDeptNumber"                  : "",
                    "RetailerID"                        : "",
                    "SKU"                               : "DROP - AMIA" if "AMIA" in shipper["recipent_line_1"] else "DROP - BAXTER",
                    "Quantity"                          : "1",
                    "UseCOD"                            : "",
                    "UseInsurance"                      : "",
                    "Saved Elements"                    : "", 
                    "Order Item Saved Elements"         : "",
                    "Carrier Notes"                     : "",
                    "shipper"                           : shipper,
                    "shipperRaw"                        : re.findall(r"SHIPPER\s([\s\S]*)\s{2}CONSIGNEE", fl.data.body)[0].strip().replace('\r',''),
                    "consignee"                         : consignee,
                    "consigneeRaw"                      : re.findall(r"CONSIGNEE\s([\s\S]*)\s{2}PICKUP LOCATION", fl.data.body)[0].strip().replace('\r',''),
                    "pickupLocation"                    : pickup_location,
                    "pickupLocationRaw"                 : re.findall(r"PICKUP LOCATION\s([\s\S]*)INSTRUCTIONS", fl.data.body)[0].strip().replace('\r',''),
                    "fullText"                          : fl.data.body
                }

                results["successess"][booking_id] = shipment_info

            except:
                failure = {
                    "ReferenceNumber"   : booking_id,
                    "shipper"           : shipper,
                    "shipperRaw"        : re.findall(r"SHIPPER\s([\s\S]*)\s{2}CONSIGNEE", fl.data.body)[0].strip().replace('\r',''),
                    "consignee"         : consignee,
                    "consigneeRaw"      : re.findall(r"CONSIGNEE\s([\s\S]*)\s{2}PICKUP LOCATION", fl.data.body)[0].strip().replace('\r',''),
                    "pickupLocation"    : pickup_location,
                    "pickupLocationRaw" : re.findall(r"PICKUP LOCATION\s([\s\S]*)INSTRUCTIONS", fl.data.body)[0].strip().replace('\r',''),
                    "fullText"          : fl.data.body
                }

                results["failures"][booking_id] = failure

    results["amiaDrops"] = len([x for x in results["successess"].values() if x.get("SKU") == "DROP - AMIA"])
    results["baxDrops"] = len(results["successess"]) - results["amiaDrops"]
    return results

if __name__ == "__main__":
    pass
    #f = sys.argv[1]
    #f = r"email.msg"
    #output_filename = datetime.now().strftime("3PLC_EI_Bookings-%Y%m%dT%H%M%S")
    #error_filename = datetime.now().strftime("3PLC_EI_Bookings-ERRORS-%Y%m%dT%H%M%S")
    #ProcessDrop(input_file=f, output_filename=output_filename)
    #input('Press any key to exit...')