{
	"name": "office365-outlook",
	"displayName": "outlook",
	"version": 1,
 	"definition": "office365/outlook/outlook.js",
	"libraries": [
	],
	"model":
	{
		"ATTACHMENT_TYPE":{"type":"attachmentType", "default":{}},
		"COERCION_TYPE":{"type":"coercionType", "default":{}},
		"searchResult" : {"type": "object", "pushToServer": "allow", "tags": {"scope": "private"}}
 	},
 	"api":
 	{
	   	"helloworld": 
	   	{
	    	"parameters":
	    	[
		    	{
					"name":"text",
					"type":"string"
				}
			]
		},
		"addRecipients":
		{
			"parameters":[
				{
					"name":"recipients",
					"type":"string[]"
				},
				{
					"name":"onError",
					"type":"function"
				}
			]
		},
		"getSelectedMessageType":
		{
			"parameters":
			[
				{
					"name":"options",
					"type":"object"
				},
				{
					"name":"callback",
					"type":"function"
				},
				{
					"name":"onError",
					"type":"function"
				}
			]
		},
		"getSelectedMessageBody":{
			"parameters":[
				{
					"name":"coercionType",
					"type":"string"
				},
				{
					"name":"options",
					"type":"object"
				},
				{
					"name":"callback",
					"type":"function"
				},
				{
					"name":"onError",
					"type":"function"
				}
			],
			"returns": "object"
		},
		"displayNewMessageForm":{
			"parameters":[
				{
					"name":"toRecipients",
					"type":"string[]"
				},
				{
					"name":"ccRecipients",
					"type":"string[]"
				},
				{
					"name":"subject",
					"type":"string"
				},
				{
					"name":"htmlBody",
					"type":"text"
				},{
					"name":"attachments",
					"type":"attachment[]"
				},{
					"name":"onError",
					"type":"function"
				}
			],
			"returns": "object"
		},
		"setBodyText":{
			"parameters": [
				{
					"name":"data",
					"type":"string"
				},
				{
					"name":"options",
					"type":"object"
				},
				{
					"name":"callback",
					"type":"function"
				},
				{
					"name":"onError",
					"type":"function"
				}
			],
			"returns": "boolean"
		}
 	},
 	"types":{
 		"coercionType":{
 			"TEXT":{"type":"string", "default":"text"},
 			"HTML":{"type":"string", "default":"html"}
 		},
 		"attachmentType":{
 			"FILE":{"type":"string", "default":"file"},
 			"ITEM":{"type":"string", "default":"item"}
 		},
 		"attachment" : {
 			 "type":{"type":"string"},
 			 "name":{"type":"string"},
 			 "media":{"type":"media"},
 			 "isInline":{"type":"boolean"},
 			 "itemId":{"type":"string"}			
 		}
 	}
}