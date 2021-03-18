# -*- coding: utf-8 -*-
"""
Created on Wed Oct 28 13:48:18 2020

@author: gretapan
"""
import teradata
import pandas as pd
import time
import feather

start = time.time()

udaExec = teradata.UdaExec(
appName="testconnections", version="1.0", logConsole=False)


with udaExec.connect(method="odbc", dsn="EBI PROD",
username={myaccount}, password={mypassword}", autocommit=True,
transactionMode="Teradata") as connect:

    query = '''SELECT


AL1.DESIGN_REGISTRATION_NUMBER AS "Design Registration",
AL1.DESIGN_ITEM_OID AS "Design Item",
AL4.MANUFACTURER_PART_NUMBER AS "Book Part",
AL1.ITEM_COMMENT AS "Item Comment",
AL1.IS_DESIGN_REGISTRATION_ACTIVE_FLAG AS "Registration Active Flag",
--
CAST(COALESCE(OPTYBRCH.SUB_RGN_MGR_USR_ID, 'UNIDENTIFIED') AS VARCHAR(21)) AS "Distributor",
AL1.DISTRIBUTOR_BRANCH_ORGANIZATION AS "Disti Branch",
OPTYBRCH.RGN_CD AS "Region",
OPTYOFFC.SUB_RGN_DESC AS "Sub Region",
OPTYOFFC.SUB_RGN_MGR_USR_ID AS "Main City Hub",
OPTYOFFC.SLS_MGR_USR_ID AS "RSM",
OPTYOFFC.ACCT_MGR_USR_ID AS "DBM",
--
AL2.END_CUSTOMER_DISPLAY_NAME AS "End Customer Name",
AL2.END_CUSTOMER_LOCATION AS "End Customer Location",
AL2.END_CUSTOMER_CITY AS "EC City",
AL2.END_CUSTOMER_STATE AS "EC State",
AL2.END_CUSTOMER_COUNTRY AS "EC Country",
EC_CG.SALES_AREA AS "EC Sales Area",
EC_CG.CBG_BUSINESS_DESCRIPTION AS "CBG",
EC_CG.SUB_CBG_BUSINESS_DESCRIPTION AS "Sub CBG",
--
AL4.ROOT_PART AS "Root Part",
AL4.PRODUCT_FAMILY_NAME AS "MAG",
MAG.MAG_MRU_DESCRIPTION AS "MAG Desc",
MAG.BUSINESS_LINE_CODE AS "BL",
MAG.BUSINESS_LINE_DESCRIPTION AS "BL Desc",
MAG.BUSINESS_UNIT_CODE AS "BU",
MAG.BUSINESS_UNIT_DESCRIPTION AS "BU Desc",

CAST(CASE WHEN CHARACTER_LENGTH(AL1.REGISTRATION_EFFORT)>0
      THEN CASE AL1.REGISTRATION_EFFORT
      WHEN 'Creation (FSL)' THEN 'DEMAND CREATION'
      WHEN 'Creation (NXP)' THEN 'EXPERT'
      WHEN 'Creation' THEN 'DEMAND CREATION'
      WHEN 'Demand Creation' THEN 'DEMAND CREATION'
      WHEN 'Shared Design (FSL)' THEN 'DEMAND CREATION'
      WHEN 'Shared Design (NXP)' THEN 'DEMAND CREATION'
      WHEN 'Shared Design Activity' THEN 'DEMAND CREATION'
      WHEN 'Shared' THEN 'DEMAND CREATION'
      WHEN 'Expert (FSL)' THEN 'EXPERT'
      WHEN 'Expert (NXP)' THEN 'EXPERT'
      WHEN 'Expert' THEN 'EXPERT'
      WHEN 'Fulfillment (FSL)' THEN 'FULFILLMENT'
      WHEN 'Base Protection' THEN 'FULFILLMENT'
      ELSE 'FULFILLMENT'
    END END AS VARCHAR(15)) AS "Reg Effort",

AL1.DESIGN_REGISTRATION_PROJECT AS "Project",
AL1.DESIGN_REGISTRATION_FUNCTION AS "Function",
--
AL1.DESIGN_REGISTRATION_STATUS AS "Design Reg Status",
AL1.DESIGN_ITEM_REGISTRATION_STATUS AS "Design Item Reg Status",
AL1.DESIGN_PART_WIN_STATUS AS "Design Part Win Status",
AL1.DESIGN_REGISTRATION_EXPIRATION_DATE AS "Expiration Date",
--
AL1.DESIGN_ITEM_CREATE_DATE AS "Item Create Date",
AL1.DESIGN_REGISTRATION_MODIFY_DATE AS "Modify Date",
REG_ITEM.LAST_MODIFY_NAME AS "Modify Name",
AL1.PROTOTYPE_DATE AS "Prototype Date",
--
AL1.DESIGN_ITEM_APPROVED_DATE AS "Approved Date",
APPROVE.FISCAL_MONTH_AS_YYYYMM AS "Approved Mth",
APPROVE.FISCAL_QUARTER_AS_YYYYQQ AS "Approved Qtr",
APPROVE.FISCAL_YEAR AS "Approved Year",
AL1.APPROVED_BY AS "Approved By",
--
AL1.PART_WIN_DATE AS "Part Win Date",
WIN.FISCAL_MONTH_AS_YYYYMM AS "Part Win Mth",
WIN.FISCAL_QUARTER_AS_YYYYQQ  AS "Part Win Qtr",
WIN.FISCAL_YEAR AS "Part Win Year",
--
AL1.DESIGN_REGISTRATION_PRODUCTION_DATE AS "Production Date",
PRODUCTION.FISCAL_MONTH_AS_YYYYMM AS "Production Mth",
PRODUCTION.FISCAL_QUARTER_AS_YYYYQQ AS "Production Qtr",
PRODUCTION.FISCAL_YEAR AS "Production Year",
--
CAST(OPTY_DFAE.DFAE_NAME AS VARCHAR(40)) AS "DFAE Name",
CAST(OPTY_SALES.SALES_NAME AS VARCHAR(40)) AS "SALES Name",
--
CURRENT_DATE AS "Run Date",
--
REG_POS_TOTAL.REG_POS_QTY AS "POS Qty",
REG_POS_TOTAL.REG_POS_AMT AS "POS Amt",
AL1.POS_AMT AS "ModelN POS Amt",
--
AL1.PROJECT_UNITS AS "Project Units",

QUOTE.AVG_DS AS "AVG DS",
QUOTE.AVG_DC AS "AVG DC",
QUOTE.AVG_SC AS "AVG SC",
QUOTE.DEBIT AS "DEBIT",
--
SUM (AL1.Value_1_Year_Production_USD) AS "DW Amt"
--
FROM 
EIM.DIST_DESIGN_REGISTRATION_FUNNEL AL1
  -- GET THE OPTY BRANCH ATTRIBUTES
  LEFT OUTER JOIN EDW.SLS_MGR_XREF OPTYBRCH ON (AL1.DISTRIBUTOR_BRANCH_SOLD_TO=OPTYBRCH.RELTNSHP_ID 
                                                AND OPTYBRCH.RELTNSHP_TYPE_TXT='FUNLOC')
--

--
  -- GET THE OPTY BRANCH OFFICE ATTRIBUTES
  LEFT OUTER JOIN EDW.SLS_MGR_XREF OPTYOFFC ON (OPTYBRCH.SLS_MGR_USR_ID=OPTYOFFC.RELTNSHP_ID 
                                                AND OPTYOFFC.RELTNSHP_TYPE_TXT='Sales Office')

  -- GET THE OPTY BRANCH MANAGER NAMES
--  LEFT OUTER JOIN EDW.SLS_MGR_XREF OPTYMGR ON (AL1.DISTRIBUTOR_BRANCH_SOLD_TO=OPTYMGR.RELTNSHP_ID 
--                                                AND OPTYMGR.RELTNSHP_TYPE_TXT='FUNLOC MGR')

  LEFT OUTER JOIN EDW.DIST_MN_END_CUSTOMER AL2 ON (AL1.END_CUSTOMER_OID = AL2.END_CUSTOMER_OID)
   LEFT OUTER JOIN EIM.CUSTOMER_GLOBAL EC_CG ON AL2.END_CUSTOMER_GLOBAL_ID = EC_CG.CUSTOMER_GLOBAL_ID
--
  LEFT OUTER JOIN EIM.SALES_PART AL4 ON (AL1.PART_OID = AL4.PART_ID)
  LEFT OUTER JOIN EIM.MANAGEMENT_ORGANIZATION_MAG_MRU MAG ON (AL4.PRODUCT_FAMILY_NAME=MAG.MAG_CODE)
  LEFT OUTER JOIN EIM.CALENDAR_FISCAL_NXP CREATE_DT ON (AL1.DESIGN_ITEM_CREATE_DATE = CREATE_DT.FISCAL_DATE)
  LEFT OUTER JOIN EIM.CALENDAR_FISCAL_NXP APPROVE ON (AL1.DESIGN_ITEM_APPROVED_DATE = APPROVE.FISCAL_DATE)
  LEFT OUTER JOIN EIM.CALENDAR_FISCAL_NXP WIN ON (AL1.PART_WIN_DATE = WIN.FISCAL_DATE)
  LEFT OUTER JOIN EIM.CALENDAR_FISCAL_NXP PRODUCTION ON (AL1.DESIGN_REGISTRATION_PRODUCTION_DATE = PRODUCTION.FISCAL_DATE)
  LEFT OUTER JOIN EIM.CALENDAR_FISCAL_NXP EXPIRE ON (AL1.DESIGN_REGISTRATION_EXPIRATION_DATE = EXPIRE.FISCAL_DATE)
  --
  LEFT OUTER JOIN EIM.DIST_DESIGN_REGISTRATION_ITEM REG_ITEM ON
   (AL1.DESIGN_ITEM_OID = REG_ITEM.DESIGN_REGISTRATION_ITEM_ID
    AND AL1.DESIGN_REGISTRATION_NUMBER = REG_ITEM.REGISTRATION_NUMBER)


  -- GET THE QUOTES
    LEFT OUTER JOIN
    (SELECT

-- CAST(QC.PRICING_DESIGN_REGISTRATION_NUMBER AS VARCHAR(15)) AS "DR_NUM",
-- QC.quote_number as "Quote",
-- QC.DEBIT_NUMBER as "Debit",
   QC.DESIGN_REGISTRATION_ITEM_ID AS "DESIGN_REGISTRATION_ITEM_ID",
   CASE WHEN QC.DEBIT_NUMBER IS NULL THEN 'N' ELSE 'Y' END AS "DEBIT",


    SUM(
   CASE WHEN QC.DISTRIBUTOR_RESALE IS NULL THEN QC.ADJUSTED_DISTRIBUTOR_RESALE/QC.EXCHANGE_RATE*QC.ITEM_QTY
   ELSE QC.DISTRIBUTOR_RESALE/QC.EXCHANGE_RATE*QC.ITEM_QTY END) /SUM(QC.ITEM_QTY) AS "AVG_DS",

    SUM(
   CASE WHEN QC.DISTRIBUTOR_COST IS NULL THEN QC.ADJUSTED_DISTRIBUTOR_COST/QC.EXCHANGE_RATE*QC.ITEM_QTY
   ELSE QC.DISTRIBUTOR_COST/QC.EXCHANGE_RATE*QC.ITEM_QTY END) /SUM(QC.ITEM_QTY) AS "AVG_DC",


   SUM(QC.MANUFACTURER_COST/QC.EXCHANGE_RATE*QC.ITEM_QTY) / SUM(QC.ITEM_QTY) AS "AVG_SC",


   ROW_NUMBER() OVER (PARTITION BY QC.DESIGN_REGISTRATION_ITEM_ID ORDER BY "DEBIT" DESC) AS SEQ

    FROM EIM.PRC_QUOTE_ITEM QC
      
--Get Latest Quote by Modified Date
--    INNER JOIN 
--     (SELECT
--      -- Get the latest Quote info
--      ST1.PRICING_DESIGN_REGISTRATION_NUMBER AS PRICING_DESIGN_REGISTRATION_NUMBER,
--      MAX(ST1.ROW_CREATE_GMT_DTTM) AS Latest_Load
--      FROM EIM.PRC_QUOTE_ITEM ST1
--      GROUP BY 1
--     ) LAST_VERSION ON QC.PRICING_DESIGN_REGISTRATION_NUMBER=LAST_VERSION.PRICING_DESIGN_REGISTRATION_NUMBER
--        AND QC.ROW_CREATE_GMT_DTTM=LAST_VERSION.Latest_Load

      INNER JOIN EIM.MATERIAL_NXP_HISTORY AL2 ON (QC.PART_12NC=AL2.PART_12NC
        AND CAST('2020-12-27' AS DATE) BETWEEN AL2.EFFECTIVE_FROM_DATE AND AL2.EFFECTIVE_TO_DATE)
      LEFT OUTER JOIN EIM.MANAGEMENT_ORGANIZATION_MAG_MRU_TV AL3 ON (AL2.MAG_CODE=AL3.MAG_CODE
            AND AL3.EFFECTIVE_DATE = CAST('2020-12-27' AS DATE))

    WHERE
     (
--    QC.Budgetary_flag='NO'
      QC.NO_BID_FLAG ='NO'
      AND QC.SPECIAL_BUY_FLAG ='NO'
   -- AND QC.PRICING_STATUS='Released'
      AND QC.workflow_status = 'Approved' 
--    AND AL3.BUSINESS_UNIT_CODE ='BU0528'
       AND QC.PRICING_DESIGN_REGISTRATION_NUMBER IS NOT NULL
       AND QC.DESIGN_REGISTRATION_ITEM_ID <> 0
     )
   GROUP BY 1,2
    having SUM(QC.ITEM_QTY) <> 0

    ) QUOTE ON
-- AL1.DESIGN_REGISTRATION_NUMBER = QUOTE.DR_NUM
    AL1.DESIGN_ITEM_OID = QUOTE.DESIGN_REGISTRATION_ITEM_ID AND QUOTE.SEQ = 1

  --
  -- GET THE REGISTERED POS $$ AMOUNT
  LEFT OUTER JOIN
    (SELECT
   POS1.DESIGN_REGISTRATION_NUMBER AS "DESIGN_REGISTRATION_NUMBER",
   POS1.DESIGN_PART_MAPPING_OID AS "DESIGN_PART_MAPPING_OID",
   SUM (CAST((POS1.DISTRIBUTOR_RESALE_USD * POS1.POS_SHIP_QTY) AS DECIMAL(18,4) ) ) AS "REG_POS_AMT",
   SUM (CAST(POS1.POS_SHIP_QTY AS DECIMAL(18) ) ) AS "REG_POS_QTY"
   FROM 
   (SELECT 
      DPP.DESIGN_REGISTRATION_NUMBER,
      DPP.DESIGN_PART_MAPPING_OID,
      DPP.DISTRIBUTOR_RESALE_USD,
      DPP.POS_SHIP_QTY,
      DPP.ITEM_12NC,
      DPP.POS_SHIP_DATE,
      DPP.REPORT_EXCLUSION_REASON
    FROM EIM.DIST_POS_PRICE DPP
    WHERE
     (
      COALESCE(SUBSTR(DPP.POS_DISTRIBUTOR_BRANCH, 1, 9), 'XXX') NOT = 'ROCHESTER'
      AND COALESCE(DPP.REPORT_EXCLUSION_REASON, 'NONE') NOT = 'XFSL MIGRATED'
      --  Exclude inventory transfer branches
      AND COALESCE(DPP.TRANSACTION_CODE, '-') NOT = 'T'
      AND COALESCE(LEFT(DPP.END_CUSTOMER_CACC,3), 'XXX') NOT = 'STA'
      AND CHARACTER_LENGTH(DPP.DESIGN_REGISTRATION_NUMBER) > 0
      AND DPP.DESIGN_REGISTRATION_NUMBER NOT = 'N/A'
      AND DPP.POS_SHIP_FISCAL_YEAR >= 2016
     )
   ) POS1

      --
   INNER JOIN EIM.CALENDAR_FISCAL_NXP NXP_DATE ON POS1.POS_SHIP_DATE = NXP_DATE.FISCAL_DATE
   INNER JOIN
    (SELECT
      FISCAL_YEAR AS Report_Yr,
      FISCAL_YEAR - 4 AS Year_Minus_1,
      FISCAL_WEEK_AS_YYYYWW AS Report_Wk
      FROM EIM.CALENDAR_FISCAL_NXP
      WHERE FISCAL_DATE = CURRENT_DATE
    ) CAL_REF 
    ON NXP_DATE.FISCAL_WEEK_AS_YYYYWW <= CAL_REF.Report_Wk
   --
   WHERE 
   ( NXP_DATE.Fiscal_Year >= CAL_REF.Year_Minus_1) 
   GROUP BY 1,2
    ) REG_POS_TOTAL ON
   AL1.DESIGN_REGISTRATION_NUMBER = REG_POS_TOTAL.DESIGN_REGISTRATION_NUMBER
   AND AL1.DESIGN_ITEM_OID = REG_POS_TOTAL.DESIGN_PART_MAPPING_OID
--

 

  -- GET THE OPTY Distributor DFAE
  LEFT OUTER JOIN 
   (SELECT
     DDR.Registration_Number,
     --
     -- Replace each occurence of two or more spaces with a single space.
     MIN(UPPER(CASE WHEN (SU.FIRST_NAME = SU.LAST_NAME)
         THEN REGEXP_REPLACE(SU.FIRST_NAME, '( ){2,}', ' ')
         ELSE REGEXP_REPLACE(SU.FIRST_NAME, '( ){2,}', ' ') || ' ' || REGEXP_REPLACE(SU.LAST_NAME, '( ){2,}', ' ')
      END)) AS "DFAE_NAME"
     FROM
       EIM.Dist_Design_Registration DDR,
       EIM.Dist_Design_Registration_User_Mapping DDRUM,
       EIM.Sales_User SU
     Where
     ((
       DDR.Design_Registration_ID = DDRUM.Design_Registration_ID
       AND DDRUM.User_ID = SU.User_ID
       AND DDR.Registration_Number NOT = 'N/A'      
       AND DDRUM.CLASSIFICATION IN ('DFAE','Engineering')
       AND SU.USER_Type = 'Distributor'
       AND SU.LAST_NAME NOT = 'NO SALES CREDIT'
     ))
    GROUP BY 1
   ) OPTY_DFAE
  ON (AL1.DESIGN_REGISTRATION_NUMBER = OPTY_DFAE.Registration_Number)
--
  -- GET THE OPTY SALES PERSON
  LEFT OUTER JOIN 
   (SELECT
     DDR.Registration_Number,
     --
     -- Replace each occurence of two or more spaces with a single space.
     MIN(UPPER(CASE WHEN (SU.FIRST_NAME = SU.LAST_NAME)
         THEN REGEXP_REPLACE(SU.FIRST_NAME, '( ){2,}', ' ')
         ELSE REGEXP_REPLACE(SU.FIRST_NAME, '( ){2,}', ' ') || ' ' || REGEXP_REPLACE(SU.LAST_NAME, '( ){2,}', ' ')
      END)) AS "SALES_NAME"
     FROM
       EIM.Dist_Design_Registration DDR,
       EIM.Dist_Design_Registration_User_Mapping DDRUM,
       EIM.Sales_User SU
     Where
     ((
       DDR.Design_Registration_ID = DDRUM.Design_Registration_ID
       AND DDRUM.User_ID = SU.User_ID
       AND DDR.Registration_Number NOT = 'N/A'      
       AND DDRUM.CLASSIFICATION IN ('DSE','DSS','Sales')
       AND SU.USER_Type = 'Distributor'
       AND SU.LAST_NAME NOT = 'NO SALES CREDIT'
     ))
    GROUP BY 1
   ) OPTY_SALES
  ON (AL1.DESIGN_REGISTRATION_NUMBER = OPTY_SALES.Registration_Number)
--
WHERE
 ((
   (APPROVE.FISCAL_YEAR >= 2019
    --OR 
    --WIN.FISCAL_YEAR >= 2018
    --OR CREATE_DT.FISCAL_YEAR >= 2019
    --OR (EXPIRE.FISCAL_YEAR >= 2019 AND AL1.DESIGN_PART_WIN_STATUS IN ('Pending','DesignWin','ProdWin'))
   )
--
   AND AL1.DESIGN_REGISTRATION_NUMBER NOT = 'N/A'
   AND AL1.TRANSITION_FLAG NOT = 'TRUE'
   -- choose mag 'RMP' for test
   AND AL4.PRODUCT_FAMILY_NAME = 'RMP'
   --AND "Distributor" IN ('ARROW','AVNET','WORLD PEACE','WT MICRO')
   AND "Design Reg Status" = 'Approved' 
   AND "Design Item Reg Status" = 'Approved'
--
--  Exclude Opty Region Japan 'SHARED' registrations - Entered only for quoting purposes
   AND NOT (OPTYBRCH.RGN_CD = 'JPN' AND COALESCE(AL1.REGISTRATION_EFFORT, '-') = 'SHARED')
-- AND MAG.BUSINESS_UNIT_CODE = 'BU0528'

 ))
--
GROUP BY
1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,
21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,
41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62
--
ORDER BY 1,2
'''

    #Reading query to df
    df = pd.read_sql(query,connect)
    # do something with df,e.g.
    print(df.head()) #to see the first 5 rows
    df.to_csv('C:/Users/{myname}/Desktop/disti_DR_project/teradataresult.csv')
    feather.write_dataframe(df, 'C:/Users/{myname}/Desktop/disti_DR_project/teradataresult.feather')
    
print(time.time() - start)
