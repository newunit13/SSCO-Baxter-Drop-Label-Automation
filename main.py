
from typing import Dict
import extract_msg
import csv
import sys
import re

f = sys.argv[1]
msg = extract_msg.Message(f)



def ProcessShipper(shipper: str) -> Dict:

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

def ProcessConsignee(consignee: str) -> Dict:

    data = [elem.strip() for elem in consignee.split('\n')]

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

    return consignee

def ProcessPickupLocation(pickup_location: str) -> Dict:

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

with open('output.tsv', 'w', newline='') as csvfile:

    fields = ["ReferenceNumber", "PurchaseOrderNumber", "ShipCarrier", "ShipService", "ShipBilling", "ShipAccount", "ShipDate", 
              "CancelDate", "Notes", "ShipTo Name", "ShipToCompany", "ShipToAddress1", "ShipToAddress2", "ShipToCity", "ShipToState", 
              "ShipToZip", "ShipToCountry", "ShipToPhone", "ShipToFax", "ShipToEmail", "ShipToCustomerID", "ShipToDeptNumber", "RetailerID", 
              "SKU", "Quantity", "UseCOD", "UseInsurance", "Saved Elements", "Order Item Saved Elements", "Carrier Notes"]

    writer = csv.DictWriter(csvfile, fieldnames=fields, delimiter='\t') 
    #writer.writeheader()


    if len(msg.attachments) > 0:
        for i, fl in enumerate(msg.attachments):

            shipper = ProcessShipper(re.findall(r"SHIPPER\s([\s\S]*)\s{2}CONSIGNEE", fl.data.body)[0].strip().replace('\r',''))
            consignee = ProcessConsignee(re.findall(r"CONSIGNEE\s([\s\S]*)\s{2}PICKUP LOCATION", fl.data.body)[0].strip().replace('\r',''))
            pickup_location = ProcessPickupLocation(re.findall(r"PICKUP LOCATION\s([\s\S]*)INSTRUCTIONS", fl.data.body)[0].strip().replace('\r',''))

            shipment_info = {
            "ReferenceNumber"                   : re.findall(r"Booking ID: (.*)", fl.data.body)[0].strip().replace('\r',''),
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
            "Carrier Notes"                     : ""
            }

            writer.writerow(shipment_info)

