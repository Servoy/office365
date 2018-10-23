{
	"name": "office365-outlook",
	"displayName": "outlook",
	"version": 1,
 	"definition": "office365/outlook/outlook.js",
	"libraries": [
	],
	"model":
	{
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
 		}
 	}
}