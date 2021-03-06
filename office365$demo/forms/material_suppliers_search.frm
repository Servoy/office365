dataSource:"db:/example_data/suppliers",
encapsulation:60,
items:[
{
anchors:11,
horizontalAlignment:0,
location:"0,0",
size:"250,60",
styleClass:"md-primary title",
text:"Supplier",
typeid:7,
uuid:"2AB9D08D-FA9F-4E3E-A71F-5FA6C89D520B"
},
{
anchors:11,
json:{
anchors:11,
location:{
x:20,
y:276
},
onActionMethodID:"3A759C9F-52B0-4317-AB13-B309B23707F2",
size:{
height:47,
width:210
},
styleClass:"md-raised md-primary",
text:"Insert supplier address",
transparent:false
},
location:"20,276",
size:"210,47",
typeName:"angularmaterial-mdbutton",
typeid:47,
uuid:"4E87123B-2FDD-4546-98F5-8C801D0DF18D"
},
{
anchors:11,
json:{
anchors:11,
location:{
x:21,
y:327
},
onActionMethodID:"6BEB1E98-176B-44FD-9F7D-2CE43106B7B1",
size:{
height:47,
width:209
},
styleClass:"md-raised md-primary",
text:"Archive text",
transparent:false
},
location:"21,327",
size:"209,47",
typeName:"angularmaterial-mdbutton",
typeid:47,
uuid:"68704FE2-2B17-402F-A57F-8E34B5016E26"
},
{
anchors:11,
json:{
anchors:11,
dataProviderID:"$supplierid",
label:"Search supplier",
location:{
x:20,
y:60
},
onDataChangeMethodID:"16AF10D9-F47B-4CEC-9CF7-C7DA94408C18",
size:{
height:50,
width:210
},
valuelist:"71FCBCFA-022D-4C17-8896-C3706A901AD2",
valuelistID:"71FCBCFA-022D-4C17-8896-C3706A901AD2"
},
location:"20,60",
name:"mdautocomplete_807",
size:"210,50",
typeName:"angularmaterial-mdautocomplete",
typeid:47,
uuid:"7D3CB7CA-92E9-4B7E-A728-BBA7988E7D1A"
},
{
anchors:3,
displaysTags:true,
location:"203,78",
name:"search",
size:"24,28",
text:"<span class=\"glyphicon glyphicon-search\"/>",
typeid:7,
uuid:"84534223-C3B4-426D-8DE2-1268556B49A5"
},
{
height:455,
partType:5,
typeid:19,
uuid:"8DAFEAAF-C3A4-4307-9C95-3875F5863575"
},
{
anchors:11,
json:{
anchors:11,
dataProviderID:"address",
enabled:true,
label:"Address",
location:{
x:20,
y:168
},
size:{
height:50,
width:210
}
},
location:"20,168",
name:"md_address",
size:"210,50",
typeName:"angularmaterial-mdinput",
typeid:47,
uuid:"A01A737A-4395-4B37-9D43-E841589C55BE"
},
{
anchors:15,
dataProviderID:"description",
displayType:1,
location:"21,382",
size:"210,64",
typeid:4,
uuid:"C82BB82F-499C-4FE0-B9DD-DA1F6D8FD03A"
},
{
anchors:11,
json:{
anchors:11,
dataProviderID:"city",
enabled:true,
label:"City",
location:{
x:20,
y:222
},
size:{
height:50,
width:210
}
},
location:"20,222",
name:"md_city",
size:"210,50",
typeName:"angularmaterial-mdinput",
typeid:47,
uuid:"D2B3FF9E-49E6-45BB-8922-FE3EB333ADBE"
},
{
anchors:11,
json:{
anchors:11,
dataProviderID:"companyname",
enabled:true,
label:"Company",
location:{
x:20,
y:114
},
size:{
height:50,
width:210
}
},
location:"20,114",
size:"210,50",
typeName:"angularmaterial-mdinput",
typeid:47,
uuid:"D4816E07-436E-4C31-AC8D-7B708F7843B3"
}
],
name:"material_suppliers_search",
showInMenu:true,
size:"250,490",
typeid:3,
uuid:"49CD60CE-9BAB-4320-AF97-E314843559DA"