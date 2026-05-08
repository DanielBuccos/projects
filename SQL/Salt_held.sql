--Begin select statement
SELECT DISTINCT VBUK.VBELN AS "Order No",
VBAK.BSTNK AS "PO No",
VBAK.ERDAT AS "created date",
VBEP.MBDAT AS "Ship Date",
VDATU AS "Req Del Date",
VBAP.WERKS AS "Plant",
VBAK.KUNNR AS " SOLD-TO NO",
Sold.NAME1 AS "SOLD-TO NAME",
VAKPA.KUNDE AS "SHIP-TO NO",
Ship.NAME1 AS "SHIP-TO NAME",
VBAP.ARKTX AS "MATERIAL",
VBKD.ZTERM AS "PAYMENT TERMS",
udmbpsegments.partner AS "PAYER",
udmbpsegments.coll_specialist AS "CREDIT CONTACT",

--Case for different credit statuses
CASE
WHEN CMGST = "A" THEN "APPROVED"
WHEN CMGST = "B" THEN "BLOCKED"
WHEN CMGST = "D" THEN "RELEASED"
END AS "CREDIT STATUS",

--Case for different billing blocks
CASE
WHEN FAKSK = "02" THEN "COMPL CONFIRM MISSING"
WHEN FAKSK = "03" THEN "PRICES INCOMPLETE"
END AS "BILLING BLOCK",

--Case for different sales order statuses
CASE
WHEN UVPRS = "A" THEN "INCOMPLETE"
WHEN UVPRS = "B" THEN "PARTIALLY INCOMPLETE"
WHEN UVPRS = "C" THEN "COMPLETE"
END AS "COMPLETION STATUS",

--Case for different delivery blocks
CASE
WHEN LIFSK = "51" THEN "PRESHIP BLOCK"
WHEN LIFSK = "54" THEN "ONLINE DELIVERY BLOCK"
END AS "DEL BLOCK",

--Case for Customer Segment
CASE
WHEN LPRIO = "28" THEN "G"
WHEN LPRIO = "29" THEN "S"
WHEN LPRIO = "30" THEN "B"
WHEN LPRIO = "31" THEN "G"
WHEN LPRIO = "32" THEN "G"
WHEN LPRIO = "33" THEN "S"
WHEN LPRIO = "34" THEN "G"
WHEN LPRIO = "35" THEN "B"
WHEN LPRIO = "36" THEN "S"
WHEN LPRIO = "37" THEN "B"
WHEN LPRIO = "38" THEN "S"
WHEN LPRIO = "39" THEN "B"
END AS "SEGMENT",

VBAK.ERNAM AS "Created by",
Ship.LAND1 AS "SHIP-TO COUNTRY CODE",
ZOCS_CUSTATTRIB.zzsales1_desc AS 'SALES PERSON',
VBKD.INCO1 AS 'INCOTERMS'

--Table selections and join statements
FROM prd_product_tc2.VBUK
JOIN prd_product_tc2.VBAK ON VBUK.VBELN = VBAK.VBELN
JOIN prd_product_tc2.VAKPA ON VAKPA.VBELN = VBAK.VBELN
JOIN prd_product_tc2.VBAP ON VBAP.VBELN = VBAK.VBELN
JOIN prd_product_tc2.VBEP ON VBAP.VBELN = VBEP.VBELN
JOIN prd_product_tc2.KNA1 AS "Sold" ON VBAK.KUNNR = Sold.KUNNR
JOIN prd_product_tc2.KNA1 AS "Ship" ON VAKPA.KUNDE = Ship.KUNNR
JOIN prd_product_tc2.VBKD ON VBUK.VBELN = VBKD.VBELN
JOIN prd_product_tc2.VBPA ON VBUK.VBELN = VBPA.VBELN
JOIN prd_product_fscm_s4.udmbpsegments ON prd_product_fscm_s4.udmbpsegments.partner = prd_product_tc2.VBPA.KUNNR
JOIN prd_product_tc2.ZOCS_CUSTATTRIB ON prd_product_tc2.ZOCS_CUSTATTRIB.zzkunn2 = Ship.KUNNR

--And statements begin here
WHERE 1=1

--Checks for only orders with a held status of some type.
AND (VBUK.CMGST IN ("B","C") OR VBAK.FAKSK IS NOT NULL OR VBAK.LIFSK IS NOT NULL OR VBUK.UVPRS IN ("A","B") OR VBUK.SPSTG = "C")

--Returns only open sales orders
AND VBAK.VKORG IN ("1001","1019")
AND VBAK.VTWEG IN ("01","03")
AND VBAK.SPART IN ("01","15")
AND VBAK.AUART = 'ZOR'
AND VBUK.LFGSK IN ("A","B")
AND VAKPA.PARVW = "WE"
AND VBAP.ABGRU IS NULL
AND VBPA.parvw = 'RG'
AND udmbpsegments.coll_segment = 'TC2'
AND udmbpsegments.valid_until = '9999-12-31'
AND prd_product_tc2.ZOCS_CUSTATTRIB.zzprctr = 'A201'
AND prd_product_tc2.ZOCS_CUSTATTRIB.zzparvw = 'WE'

--Sort order
ORDER BY VBAK.KUNNR, VAKPA.KUNDE ,VBEP.MBDAT
