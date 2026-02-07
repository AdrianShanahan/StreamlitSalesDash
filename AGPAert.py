import pyodbc
import pandas as pd
import win32com.client as win32

# SQL Server Connection
conn_str = (
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=corp-bi;"
    "Database=jetaxdwh;"
    "Trusted_Connection=yes;"
    "Encrypt=yes;TrustServerCertificate=yes;"
)

conn = pyodbc.connect(conn_str)

# AGP Query
sql = """
WITH BaseItems AS (
SELECT
      a.[Item No]
	  ,REPLACE(a.[item Search Name],CONCAT(a.[Product Status],' - '),'') AS 'Search Name'
	  ,REPLACE(REPLACE(REPLACE(a.[item Search Name],CONCAT(a.[Product Status],' - '),''),'R','*'),'L','*') AS 'Matching Name'
	  ,CASE
			WHEN a.[Item Search Name] LIKE '%LA2%' OR a.[Item Search Name] LIKE '%LB2%' THEN 'L'
			WHEN a.[Item Search Name] LIKE '%RA2%' OR a.[Item Search Name] LIKE '%RB2%' THEN 'R'
			ELSE NULL
	  END AS 'L/R'
	  ,CASE
			WHEN a.[Production Type] = 'BOM' AND a.[Item Price Tol Group] = 'SPA' THEN 'BOTH'
			WHEN a.[Production Type] = 'BOM' AND a.[Item Price Tol Group] !='SPA' THEN 'BOM'
			WHEN a.[Production Type] != 'BOM' AND a.[Item Price Tol Group] = 'SPA' THEN 'SPA'
			WHEN a.[Production Type] != 'BOM' AND a.[Item Price Tol Group] != 'SPA' THEN 'Neither'
	  END AS 'BOM/SPA'
      ,a.[Default Order Type]
      ,a.[Product Status]
  FROM [JetAxDwh].[001].[Item] a
  WHERE a.[S-Line Disc Group] = 'S-AGP' 
  AND a.[Exclude From D365] IS NULL
  AND a.[Item Search Name] NOT LIKE '%AGP25%'
  ),
  L_items AS (
  SELECT * FROM BaseItems WHERE [L/R] = 'L'
  ),
  R_Items AS (
  SELECT * FROM BaseItems WHERE [L/R] = 'R'
  ),
  STAs AS (
  SELECT
    [Item No]
    ,[Amount]
    ,[Date From]
  FROM [JetAxDwh].[001].[Trade Agreements] a
  WHERE a.[Is Valid Today] = '1'
  AND a.[Relation ID]= '4'
  )

SELECT DISTINCT
	L.[Item No]
	,L.[Search Name]
	,L.[L/R]
	,L.[BOM/SPA]
	,L.[Default Order Type]
	,L.[Product Status] AS 'Status'
	,L_STA.[Amount] AS 'L STA Amount'
	,CAST(L_STA.[Date From] AS DATE) AS 'L STA Date'
	,R.[Item No] AS 'Item No2'
	,R.[Search Name] AS 'Search Name2'
	,R.[L/R] AS 'L/R2'
	,R.[BOM/SPA] AS 'BOM/SPA 2'
	,R.[Default Order Type] AS 'Default Order Type2'
	,R.[Product Status] AS 'Status2'
	,R_STA.[Amount] AS 'R STA Amount'
	,CAST(R_STA.[Date From] AS DATE) AS 'R STA Date'
	,CASE
			WHEN L_STA.[Amount] = R_STA.[Amount] AND (L_STA.[Date From] > DATEADD(day,-30,CURRENT_TIMESTAMP) OR R_STA.[Date From] > DATEADD(day,-30,CURRENT_TIMESTAMP)) THEN 'Already updated'
			WHEN L_STA.[Amount] = R_STA.[Amount] AND (L_STA.[Date From]  <= DATEADD(day,-30,CURRENT_TIMESTAMP) AND R_STA.[Date From] <= DATEADD(day,-30,CURRENT_TIMESTAMP)) THEN 'Prices Match, older dates'
			WHEN L.[Product Status] IN ('OBS','OWD') AND R.[Product Status] IN ('OBS','OWD') THEN 'Ignore'
			WHEN L.[Default Order Type] = 'purchase order' AND R.[Default Order Type] = 'production' AND L_STA.[Amount] != R_STA.[Amount] THEN 'Update R to Match L'
			WHEN L.[Default Order Type] = 'production' AND R.[Default Order Type] = 'purchase order' AND L_STA.[Amount] != R_STA.[Amount] THEN 'Update L to Match R'
			WHEN L.[Default Order Type] = 'purchase order' AND R.[Default Order Type] = 'purchase order' AND (L.[Product Status] = R.[Product Status]) THEN 'Update both with Quote'
			WHEN L.[Product Status] != R.[Product Status] THEN 'Check Statuses'
	END AS 'Notes'
FROM L_items L
	LEFT OUTER JOIN R_Items R ON L.[Matching Name] = R.[Matching Name]
	LEFT JOIN STAs L_STA ON L.[Item No] = L_STA.[Item No]
	LEFT JOIN STAs R_STA ON R.[Item No] = R_STA.[Item No]
WHERE ((L.[Product Status] = '' AND R.[Product Status] = '') OR L.[Product Status] != R.[Product Status])
ORDER BY 
	CASE
			WHEN L_STA.[Amount] = R_STA.[Amount] AND (L_STA.[Date From] > DATEADD(day,-30,CURRENT_TIMESTAMP) OR R_STA.[Date From] > DATEADD(day,-30,CURRENT_TIMESTAMP)) THEN 'Already updated'
			WHEN L_STA.[Amount] = R_STA.[Amount] AND (L_STA.[Date From]  <= DATEADD(day,-30,CURRENT_TIMESTAMP) AND R_STA.[Date From] <= DATEADD(day,-30,CURRENT_TIMESTAMP)) THEN 'Prices Match, older dates'
			WHEN L.[Product Status] IN ('OBS','OWD') AND R.[Product Status] IN ('OBS','OWD') THEN 'Ignore'
			WHEN L.[Default Order Type] = 'purchase order' AND R.[Default Order Type] = 'production' AND L_STA.[Amount] != R_STA.[Amount] THEN 'Update R to Match L'
			WHEN L.[Default Order Type] = 'production' AND R.[Default Order Type] = 'purchase order' AND L_STA.[Amount] != R_STA.[Amount] THEN 'Update L to Match R'
			WHEN L.[Default Order Type] = 'purchase order' AND R.[Default Order Type] = 'purchase order' AND (L.[Product Status] = R.[Product Status]) THEN 'Update both with Quote'
			WHEN L.[Product Status] != R.[Product Status] THEN 'Check Statuses'
	END DESC
"""

df = pd.read_sql(sql, conn)

wanted = {"Check Statuses", "Update L to Match R", "Update R to Match L", "Update both with Quote"}
filtered_df = df[df["Notes"].isin(wanted)].copy()

# HTML for Email
html_table = filtered_df.to_html(index=False)
html_content = f"""
<html>

  <body>
    <p>Please find below the AGP items that require price updates:</p>
    {html_table}
    </body>
</html>
"""
preview_path = r"C:\Users\ashanahan\Downloads\Python\EmailAlerts\Previews\AGPEmailPreview.html"
with open(preview_path, "w") as file:
    file.write(html_content)

print(f"HTML preview saved to {preview_path}")

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.To = 'Pricing@forceamerica.com'
mail.Subject = 'AGP Items Needing Price Updates'
mail.HTMLBody = html_content

if not filtered_df.empty:
      mail.SentOnBehalfOfName = "Pricing@forceamerica.com"
      mail.Send()
else:
      print("No items require price updates. Email not sent.")