library(DBI)
library(odbc)

con <- dbConnect(
  odbc::odbc(),
  Driver   = "snowflake",
  Server   = "SPGLOBALXPRESSCLOUD-XF_READER_MAGMUTUALINSURANCE.snowflakecomputing.com",
  UID      = "Dontay_harris",
  PWD      = Sys.getenv("CONNECTION_SP_DATA_DATABASE_PASSWORD"),
  Database = "MI_XPRESSCLOUD",
  Warehouse= "XF_READER_MAGMUTUALINSURANCE_WH",
  Schema   = "XPRESSFEED"
)


 qry <- '
SELECT top 80000
  CI1.ENTITYNAME_4 AS "Entity Name ",
  PC4.*
FROM SNL_PC_SCHEDULED_SD0004 PC4
JOIN XPRESSFEED.SNL_PC_INSURANCEPRODUCTFILINGS_CORPORATECOMPANYINFORMATION CI1
  ON PC4.SNLSTATUTORYENTITYKEY_1 = CI1.SNLSTATUTORYENTITYKEY_14
WHERE SUBSTR(PC4.SNLDATASOURCEPERIOD_3, 5, 1) = \'Q\' and SUBSTR(PC4.SNLDATASOURCEPERIOD_3, 1, 4) in (\'2024\',\'2023\',\'2025\')
'
data <- dbGetQuery(con, qry)


View(data)
saveRDS(data,"Snowflake.rds")


dbDisconnect(con)
