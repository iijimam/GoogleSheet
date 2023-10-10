Class GoogleSheet.SheetTemplate Extends JSONTemplate.Base
{

Property SheetData As list Of GoogleSheet.SheetTemplate;

XData Template [ MimeType = application/json ]
{
{
  "valueInputOption": "RAW",
  "data": ["#(..SheetData)#"],
  "includeValuesInResponse": true,
  "responseValueRenderOption": "FORMATTED_VALUE"
}
}

ClassMethod CreateJSON() As %DynamicObject
{
    set o1=##class(GoogleSheet.DataTemplate).%New()
    set o1.RangeName="sales!A1:C4"
    set ar1=[]
    do ar1.%Push("2022-12-23")
    do ar1.%Push(200)
    do ar1.%Push(299)
    set ar2=[]
    do ar2.%Push("2023-2-3")
    do ar2.%Push(190)
    do ar2.%Push(39)
    do o1.Data.Insert(ar1)
    do o1.Data.Insert(ar2)

    set o2=##class(GoogleSheet.DataTemplate).%New()
    set o2.RangeName="sales2!A1:C3"
    set ar3=[]
    do ar3.%Push("2023-10-2")
    do ar3.%Push(2)
    do ar3.%Push(9)
    set ar4=[]
    do ar4.%Push("2023-10-4")
    do ar4.%Push(119)
    do ar4.%Push(59)
    do o2.Data.Insert(ar3)
    do o2.Data.Insert(ar4)

    set s=##class(GoogleSheet.SheetTemplate).%New()
    do s.SheetData.Insert(o1)
    do s.SheetData.Insert(o2)

    //do s.OutputToDevice()
    set st=s.OutputToDynamicObject(.obj)
    return obj
}

ClassMethod CreateJSON2() As %DynamicObject
{
    set o1=##class(GoogleSheet.DataTemplate).%New()
    set o1.RangeName="sales!A1:C4"
    set sub=""
    for {
      set sub=$order(^SheetData("sales",sub),1,data)
      if sub="" quit
      set array=[]
      for i=1:1:$listlength(data) {
        do array.%Push($listget(data,i))
      }
      do o1.Data.Insert(array)
      kill data,array
    }

    set o2=##class(GoogleSheet.DataTemplate).%New()
    set o2.RangeName="sales2!A1:C3"

    set sql="select SalesDate,Apple,Grape from GoogleSheet.Tbl"
    set stmt=##class(%SQL.Statement).%New()
    set stmt.%SelectMode=1
    set qStatus=stmt.%Prepare(sql)
    set rset=stmt.%Execute()
    while rset.%Next() {
      set array=[]
      do array.%Push(rset.%Get("SalesDate"))
      do array.%Push(rset.%Get("Apple"))
      do array.%Push(rset.%Get("Grape"))
      do o2.Data.Insert(array)
    }

    set s=##class(GoogleSheet.SheetTemplate).%New()
    do s.SheetData.Insert(o1)
    do s.SheetData.Insert(o2)

    //do s.OutputToDevice()
    set st=s.OutputToDynamicObject(.obj)
    return obj
}

ClassMethod CreateDummyData()
{
  kill ^SheetData
  &sql(drop table GoogleSheet.Tbl)

  set ^SheetData("sales",1)=$LISTBUILD("年月日","りんご","ぶどう")
  set ^SheetData("sales",2)=$LISTBUILD("2023年10月3日",100,299)
  set ^SheetData("sales",3)=$LISTBUILD("2023年10月4日",50,99)
  set ^SheetData("sales",4)=$LISTBUILD("2023年10月8日",6,112)

  #SQLCompile Select = odbc
  &sql(CREATE TABLE GoogleSheet.Tbl (SalesDate DATE,Apple INTEGER,Grape INTEGER))
  &sql(INSERT INTO GoogleSheet.Tbl (SalesDate,Apple,Grape) VALUES('2022-10-02',100,200))
  &sql(INSERT INTO GoogleSheet.Tbl (SalesDate,Apple,Grape) VALUES('2022-10-04',9,19))
  &sql(INSERT INTO GoogleSheet.Tbl (SalesDate,Apple,Grape) VALUES('2022-10-08',20,10))
}

}
