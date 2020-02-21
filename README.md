# AssetControl
An open source system utilizing PoweShell and SQL to track assets.  ReadMe.txt for instructions on porting to your company.

Please follow the GNU Opensource liscensing guidelines to use, modify, share, and enjoy.
Author = 'Protheophage'

For all files replace all instances of the below with your information
    <Your_Server><Your_PSRepo>
    <Your_PSRepo>
    OU=<YourOU>,DC=<YourDomain>,DC=com
    SQLSERVER:<Your_SQL_Server>Databases\Assets\Tables
    $DeptOU
    Author = 'Protheophage'
    CompanyName = 'OpenSourceTooling OST'
    <path-to-AssetTagExe>
        ##This is the path to your shared drive copy of the MS Surface Asset Tag Utility
        ###directing to a download page and executing through psexec does not work as well as copying from a domain location
    <Path-to-DymoLabel-Template>
        ##This is the path to your shared drive copy of the Dymo Label Template
    <Path-to-DymoLabelMsi>
        ##This is the path to your shared drive copy of "Dymo Label.msi"
    <Path-to-shared-dymoPrinter>


When creating your databases please structure and name them as follows:
    If you use a different structure or names please ensure you update all documents to match

Database: Assets

DBO: AssetList
    date_added
    ,date_updated
    ,asset_name
    ,asset_id
    ,asset_type_name
    ,serial_number
    ,manufacturer
    ,model
    ,description
    ,product_key
    ,status
    ,purch_price

DBO: Retired
    date_registered
    ,date_retired
    ,asset_id
    ,asset_type_name
    ,serial_number
    ,manufacturer
    ,model
    ,purch_price
    ,retired_value

DBO: Laps_Log
	date_logged
	,asset_name
	,asset_id
	,serial_number
	,laps_pw
