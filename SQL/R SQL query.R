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


qry <- 'SELECT * FROM "MI_XPRESSCLOUD"."XPRESSFEED"."SNL_PC_COREFINL_CF0008" LIMIT 10;'
data <- dbGetQuery(con, qry)


View(data)
saveRDS(data,"Snowflake.rds")

dbDisconnect(con)
