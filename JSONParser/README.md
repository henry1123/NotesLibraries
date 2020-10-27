#JSON Parser for LotusScript

## How to use :

in declare sectio ```use "JSON"```

Create new JsonParser Object
Use that object to parsing JSON text.
if return Key-Value , use  GetItem(Key) get Item Value
if return Array , use Items to get the object array , and use ```Forall item in Items``` to get the arrat content


## Sample Code:
```vba
Option Public
Option Declare

Use "JSON"

Sub Initialize
	
	Dim sJSON As String
	Dim jsonParser As New JSONParser
	
	sJSON=GetSampleJSONString
	
	Dim vResults As Variant
	Set vResults = jsonParser.Parse(sJSON)

	Dim Doc_List
	Set Doc_List=vResults.GetItem("Doc_List")

	Dim Result
	Result=""
	ForAll item In Doc_List.Items
		Dim R
		R=""
		R=R+SP1+item.GetItem("Agent_Name")
		R=R+SP1+item.GetItem("Inspect_Date")
		R=R+SP1+item.GetItem("Inspection_Person")
		'....
		
				
		Dim Check_List, R1
		R1=""
		Set Check_List=item.GetItem("Check_List")		
		ForAll citem In Check_List.Items
			Dim R2
			R2=R2+" * "+citem.GetItem("Item_name")
			R2=R2+" * "+citem.GetItem("Status")
			
			R1=R1+" -> "+R2
		End ForAll
		
		R=R+SP1+R1
		
		Result=ListConcat(Result, R)	
	End ForAll
	Result=ListTrim(Result)	

End Sub

Function GetSampleJSONString
	GetSampleJSONString=|{
"DocNo": "AGC-C-2018100427",
"Doc_List": [
{
"Agent_Name": "A公司",
"Check_List": [{
"Item_name": "標籤符合材料規範標籤規定",
"Status": "OK"
},
{
"Item_name": "外觀標籤與標示內容物一致",
"Status": "OK"
},
{
"Item_name": "外包裝無破損+/+無髒污等異常",
"Status": "OK"
},
{
"Item_name": "瓶(桶)身無洩漏/無變形/無破損+/+無髒污等異常",
"Status": "OK"
}
],
"Inspect_Date": "2018/10/09",
"Inspection_Person": "Joyce+Chuang",
"Lot_No": "180821P",
"Remark": "",
"Result": "OK",
"Supplier_Name": "AAA"
},
{
"Agent_Name": "A公司",
"Check_List": [{
"Item_name": "標籤符合材料規範標籤規定",
"Status": ""
},
{
"Item_name": "外觀標籤與標示內容物一致",
"Status": ""
},
{
"Item_name": "外包裝無破損+/+無髒污等異常",
"Status": ""
},
{
"Item_name": "瓶(桶)身無洩漏/無變形/無破損+/+無髒污等異常",
"Status": ""
}
],
"Inspect_Date": "2018/10/09",
"Inspection_Person": "Joyce+Chuang",
"Lot_No": "180824W",
"Remark": "",
"Result": "OK",
"Supplier_Name": "AAA"
}
]
}|
End Function
```

