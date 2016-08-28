## Journal Entry Completeness Routine v1.1.0

This is update to initial release of the *Journal Entry Completeness Routine*.
- Version: *v1.1.0*
- Developed by *Shahbaz Khan*
- Initial release date: *March 26, 2016*
- Updated release date: *August 29, 2016*

#### Background
Journal Entry testing has been key process during any financial audit due to risk of fraudulent financial reporting as per International Standards of Auditing.

Due to extensive on going modernization of technology, ERP and highly sophisticated software packages such as *Oracle, SAP, Microsoft Dynamics* have been used in production environments for effective reporting and efficient management of financial data and specially the core transactional data with great details of accounting and business attributes.  

Therefore, journal entries for any financial period have significant amount of data that is required to be covered in audit risk universe. Auditors are now using CAATs such as **ACL, Caseware IDEA** for testing such huge transactional data more precisely by focusing on more riskiest area within the populations.

Before applying automated audit tests, transaction data is required to be imported in the CAAT. The first question arises to mind after import process is particularly about the completeness of the data imported from the source file also from the source system. After verifying completeness of transactional data, auditor can proceed with automated audit tests to meet its audit objective.

#### Purpose
The purpose this routine is to automate process of verifying completeness of journal entries data.

#### Compatible CAAT
This routines is made for used with **Caseware IDEA**. Routine scripting language is *IDEAScript*.

#### Mechanism

##### *Key Inputs from User:*
1. Trial balance file for opening balance
    1. Identification of opening balance field.
    2. Matching key field *i.e. Account No. in Chart of Account*.
2. Trial balance file for closing balance
    1. Identification of closing balance field
    2. Matching key field *i.e. Account No. in Chart of Account*.
3. Journal Entries file
    1. Identification of debit amount field
    2. Identification of credit amount field
    3. Matching key field *i.e. Account No. in Chart of Account*.

##### *Process:*
The routine simply takes opening balance of each account from opening trial balance file and add summarized activity in the respective account to get derived closing balance. Then it will compare derived closing balance with actual closing balance of respective account in the closing trial balance to see any differences. If Journal entries are complete, the differences would be nil.

###### *Technical Process in Caseware IDEA*
The routine simply performs the following key tasks:
* Using Summarization task the journal entries file by its key field i.e. *ACCOUNT_ID*
* Using Visual Connector task in analysis tab to join all three input files based on key field i.e. *ACCOUNT_ID*
* Using Append Field, it will append two fields with following criteria
  * Field 1 name: ```DERIVED_CLOSING```
  * Criteria 1: ``` DERIVED_CLOSING = OPENING_BAL + Summarized_JE_DR - Summarized_JE_CR ```

  * Field 2 name: ```DIFFERENCE```
  * Criteria 2: ``` DERIVED_CLOSING - CLOSING_BAL```

##### _Additional Features & Fixes_

*Change log - 1.1.0*
* Fix: Join relation changed for Visual Connector task in *RelateDatabase* function. Now ```id0``` (i.e. Closing trail balance) would be the primary database for ```id1``` (i.e. Opening trial balance) and ```id2``` (i.e. Journal entries).
* Added: ```ListBox$7()``` in dialog to capture account description in ```accdesc``` variable.
* Added: Condition in *Case 2* with Closing Trial Balance database to include only non-numeric fields in ```ListBox$7()```.
* Fix: ```IncludeAllFields``` method and used ```AddFieldtoInclude``` include to only following fields in joined database.
  - ```[ACCOUNT_ID]```
  - ```[ACCOUNT_DESCRIPTION]```
  - ```[OPENING_BAL]```
  - ```[Summarized_JE_Debit_balance]```
  - ```[Summarized_JE_Credit_balance]```
  - ```[CLOSING_BAL]```
* Added: Error handler for databases that does not have any account description fields in closing trial balance database.

*Change Log - 1.0.3*
* Fix: ```Client.CloseAll``` from Append1 function to Sub Main due to error in IDEA v9
* Fix: ```Client.OpenDatabase``` from Append1 function to Sub Main due to error in IDEA v9

*Change Log - 1.0.2*
* Added: error handling routines that verifies dialog menu before processing with list of 6 errors.
* Added: Unique file name
* Added: Add separately the summarized database with name "Summarize by Account No."
* Added: Used virtual field type for while appending that gives ability to see and change formula.
* Added: Underlying speedup keys for fast input i.e. press ```ALT``` key shows underlines letters for faster input.
* Fix: Grouped all three files and format changes for clear visualization

##### _Upcoming Features_
* Help button for displaying help while giving error
* Data preparation routine for this routine for datasets that have additional fields or separate debit and credit fields.
* Report generation into word document with details.
