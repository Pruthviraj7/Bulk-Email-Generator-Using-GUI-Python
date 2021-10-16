import pandas as pd
data=pd.read_excel("F:\GUI Python\Project\Bulk_Email\sample.xlsx")

if "Email" in data.columns:
    emails=list(data['Email'])
    c=[]
    for i in emails:
        if pd.isnull(i)==False:
            c.append(i)
    emails=c
    print(emails)
else:
    print("Not exist")

# wadefrtdxrytdfbm

 