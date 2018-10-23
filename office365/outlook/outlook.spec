{
	"name": "office365-outlook",
	"displayName": "outlook",
	"version": 1,
 	"definition": "office365/outlook/outlook.js",
	"libraries": [
	],
	"model":
	{
		"searchResult" : {"type": "object", "pushToServer": "allow", "tags": {"scope": "private"}}
 	},
 	"api":
 	{
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
 	}
}