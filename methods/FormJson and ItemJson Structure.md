FormJson example
---============== Form ==============
--------- "id": 61075
--------- "name": Testing Ground #1
--------- "studyEventName": SCREENING
------ "cohort":
------------ "id": null
------------ "name": null
--------- "epoch":
--------------- "id": null
--------------- "name": null
--------- "timepoint": null
--------- "canceled": false
--------- "dataCollectionStatus": Nonconformant
------ "subject":
------------ "id": 1387
------------ "locked": false
------------ "screeningNumber": S0003
------------ "leadInNumber": null
------------ "randomizationNumber": null
------------ "subjectStudyStatus": Active
------------ "subjectEligibilityType": Unspecified
--------- "volunteer":
--------------- "id": 3
--------------- "initials": L-A
--------------- "age": 52
--------------- "sexMale": true
--------------- "dateOfBirth": 1973-02-28
------------ "site":
------------------ "name": CenExel ACT
------------------ "volunteerRegionLabel": State
------------------ "heightUnits": Centimeters
------------------ "weightUnits": Kilo
------------------ "address": 2441 W La Palma Ave, #140
------------------ "city": Anaheim
------------------ "stateProvince": CA
------------------ "postalCode": 92801
------------------ "country": USA
------------------ "phoneNumber": null
------------------ "faxNumber": null
------------------ "timeZoneId": America/Los_Angeles
------------------ "volunteerDisplayType": FullDobSexInitials
------------------ "volunteerEnrollmentConfirmationStatement": null


== ITEM GROUPS ==

---========== [Form: Testing Ground #1] [ItemGroup 0]: Subject Info ==========
------============== ItemGroup[0] ==============
------------ "id": 163609
------------ "collectionTimems": 1759507427000
------------ "deviationSeconds": null
------------ "itemGroupRepeatKey": 1
------------ "name": Subject Info
------------ "domain": null
------------ "dataCollectionStatus": Complete
------------ "canceled": false

------- [Form: Testing Ground #1] [ItemGroup: Subject Info] Item 0: ID ----
---------============== Item ==============
--------------- "id": 788332
--------------- "name": ID
--------------- "dataType": text
--------------- "dataCollectionStatus": Complete
--------------- "sasFieldName": null
------------ "codeListItems":
------------ "measurementUnits":
--------------- "measurementUnit": null
--------------- "value": 1387
--------------- "outOfRange": false
--------------- "nonconformantMessage": null
--------------- "length": 10
--------------- "significantDigits": null
--------------- "canceled": false

============== ITEM JSON DATA ==============
id: 788352
name: Current Date
dataType: datetime
dataCollectionStatus: Complete
sasFieldName: null
codeListItems: []
measurementUnits: []
measurementUnit: null
value: 2025-10-03T11:05:06
outOfRange: false
nonconformantMessage: null
length: 10
significantDigits: null
canceled: false
dateValueMs: 1759514706000
============== END ITEM JSON DATA ==============


Subject Values:
subject.id: 1387
subject.locked: false
subject.screeningNumber: S0003
subject.leadInNumber: null
subject.randomizationNumber: null
subject.subjectStudyStatus: Active
subject.subjectEligibilityType: Unspecified
subject.volunteer.id: 3
subject.volunteer.initials: L-A
subject.volunteer.age: 52
subject.volunteer.sexMale: true
subject.volunteer.dateOfBirth: 1973-02-28
subject.volunteer.site.name: CenExel ACT
subject.volunteer.site.volunteerRegionLabel: State
subject.volunteer.site.heightUnits: Centimeters
subject.volunteer.site.weightUnits: Kilo
subject.volunteer.site.address: 2441 W La Palma Ave, #140
subject.volunteer.site.city: Anaheim
subject.volunteer.site.stateProvince: CA
subject.volunteer.site.postalCode: 92801
subject.volunteer.site.country: USA
subject.volunteer.site.phoneNumber: null
subject.volunteer.site.faxNumber: null
subject.volunteer.site.timeZoneId: America/Los_Angeles
subject.volunteer.site.volunteerDisplayType: FullDobSexInitials
subject.volunteer.site.volunteerEnrollmentConfirmationStatement: null


{
    "item": {
        "id": 136424,                      // Unique ID for the item data
        "name": "AE_START_DATE",           // Name of the item (e.g., 'Adverse Event Start Date')
        "dataType": "incompleteDatetime",  // Data type (e.g., 'string', 'date', 'incompleteDatetime')
        "dataCollectionStatus": "Complete",// Status: 'Complete', 'In Progress', etc.
        "sasFieldName": "AESTDTC",         // Corresponding SAS field name
        "value": "2024-08-05T13:39:45",      // Captured value (can be incomplete)
        "codeListItems": [                 // Optional: List of allowed coded values
            {
                "codedValue": "CS",         // Coded value
                "decode": "Clinically Significant" // Meaning of coded value
            },
            {
                "codedValue": "NCS",
                "decode": "Not Clinically Significant"
            }
        ],
        "measurementUnits": [],        // Array of applicable measurement units
        "measurementUnit": null,       // Specific measurement unit (if used)
        "outOfRange": true,            // Whether the value is out of an acceptable range
        "nonconformantMessage": null,  // Message if the item is nonconformant
        "length": null,                // Optional: Length constraint for the value
        "significantDigits": null,     // Optional: Number of significant digits allowed
        "canceled": false              // Whether this item has been canceled
    }
}

Data type format:
datetime: 2026-02-05T15:32:12
partial datetime: 2026-02-17T08:40:21
partial date: 2026-02-17
partial time: 08:40:21
incomplete datetime: 2026-02-17
incomplete time: 08:40:22
time: 08:40:22
date: 2026-02-17
