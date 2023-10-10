Class GoogleSheet.DataTemplate Extends JSONTemplate.Base
{

XData Template [ MimeType = application/json ]
{
{
    "range": "#(..RangeName)#",
    "majorDimension": "ROWS",
    "values": ["#(..Data)#"]
}
}

Property RangeName As %String;

Property Data As list Of %DynamicArray;

}
