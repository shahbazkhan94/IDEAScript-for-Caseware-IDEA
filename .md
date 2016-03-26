## Journal Entry Completeness Routine v1.0

This is initial release of the *Journal Entry Completeness Routine*. 
- Version: *v1.0*
- Authored by *Shahbaz Khan*
- Dated: *March 26, 2016*

#### Background
Journal Entry testing has been key process during any financial audit due to risk of fraudulent financial reporting as per International Standards of Auditing. 

Due to extensive on going modernization of technology, ERP and highly sophisticated softwares packages  such as *Oracle, SAP, Microsoft Dynamics* are being used in production environments for management of financial data and specially the core transaction data with its no. of details.

Therefore, journal entries for any financial period have significant amount of data that is required to be covered in audit risk universe. Auditors are now using CAATs such as **ACL, Caseware IDEA** for testing such huge transnational data and more precisely focusing on more riskiest area.

Before applying automated audit tests, transaction data is required to be imported in the CAAT. The first question arises to mind after import process is particularly about the completeness of the data imported from the source file also from the source system. After verifying completeness of transaction data, auditor can proceed with automated audit tests to meet its audit objective.

#### Compatible CAAT
This routines is made for used with **Caseware IDEA**. Routine scripting language is *IDEAScript*.

#### Purpose
The purpose this routine is to automate process of verifying completeness of journal entries data.

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
The routine simply takes opening balance from opening trial balance file and add calculated activity in a particular particular account to get derived closing balance. Then it will compare derived closing balance with actual closing balance in the closing trial balance to see any differences. If Journal entries are complete, the differences will be nil.

###### *Technical Process in Caseware IDEA*
The routine simply performs the following key tasks:
* Using Summarization task the journal entries file by its key field i.e *ACCOUNT_ID*
* Using Visual Connector task in analysis tab to join all three input files based on key field i.e. *ACCOUNT_ID
* Using Append Field, it will append two fields with following criteria
  * Field 1 name: ```DERIVED_CLOSING```
  * Criteria 1: ``` DERIVED_CLOSING = OPENING_BAL + Summarized_JE_DR - Summarized_JE_CR ```

  * Field 2 name: ```DIFFERENCE```
  * Criteria 2: ``` DERIVED_CLOSING - CLOSING_BAL```

##### _Additional Features_
* Added error handling routines that verifies dialog menu before processing with list of 6 errors.
* Unique file name
* Add separately the summarized database with name "Summarize by Account No." 
* Used virtual field type for while appending that gives ability to see and change formula.

##### _Upcoming Features_
* Help button for displaying help while giving error
* Underlying word shortcuts for fast input
* Data preparation routine for this routine for datasets that have additional fields or separate debit and credit fields.
* Report generation into word document with details.
