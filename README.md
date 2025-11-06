curl -X POST http://localhost:8080/demand-letters-api \
  -H "Content-Type: application/json" \
  -d '{
    "format": "pdf",
    "template": "DL1",
    "data": {
      "our_ref": "STIMA/REC/2025/00123",
      "date": "2025-11-05",
      "customer": {
        "name": "John Doe",
        "account_number": "L0012142",
        "address_line_1": "P.O. Box 12345-00100",
        "address_line_2": "Nairobi"
      },
      "loan": {
        "principal_amount": "304,538.58",
        "arrears_amount": "62,781.20",
        "interest_rate": "13%"
      },
      "guarantors": [
        { "name": "ABDULKADIR HAMISI BADI", "address": "P. O BOX 306-80300 Voi" },
        { "name": "NOEL MUNYAE MAKUKWI",   "address": "P. O BOX 13828-00800 Westlands" }
      ]
    }
  }'

curl -X POST http://localhost:8080/letters \
  -H "Content-Type: application/json" \
  -d '{
    "template_code": "DL2",
    "format": "pdf",
    "data": {
      "our_ref": "STIMA/REC/2025/00456",
      "date": "2025-11-05",
      "customer": { "name": "Jane A. Doe", "account_number": "L0098765" },
      "loan": { "principal_amount": "250,000.00", "arrears_amount": "45,000.00" },
      "guarantors": [{ "name": "John G", "address": "P.O. Box 123 Nairobi" }]
    }
  }' --output demand2.pdf

