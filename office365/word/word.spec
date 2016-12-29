{
	"name": "office365-word",
	"displayName": "word",
	"version": 1,
 	"definition": "office365/word/word.js",
	"libraries": [
	],
	"model":
	{
		"searchResult" : {"type": "object", "pushToServer": "allow", "tags": {"scope": "private"}}
 	},
 	"api":
 	{
 	 	"isOfficeEnabled": {
	    	"parameters": [],
			"returns" : "boolean"
		},
		"getDownloadUrl": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"getDocumentUrl": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"getSelectedData": {
	    	"parameters": [{
					"name": "type",
					"type": "string",
					"optional" : true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "object"
				
		},
		"setSelectedData": {
	    	"parameters": [{
					"name": "text",
					"type": "object"
				},{
					"name": "type",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"getBodyText": {
	    	"parameters": [{
					"name": "loadOptions",
					"type": "object",
					"optional" : true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		 "getBodyHtml": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		 "getBodyOoxml": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"insertTextToBody": {
	    	"parameters": [{
					"name": "text",
					"type": "string"
				},{
					"name": "insertLocation",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"insertHtmlToBody": {
	    	"parameters": [{
					"name": "htmlText",
					"type": "string"
				},{
					"name": "insertLocation",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"insertOoxmlToBody": {
	    	"parameters": [{
					"name": "ooxmlText",
					"type": "string"
				},{
					"name": "insertLocation",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"selectBody": {
	    	"parameters": [{
					"name": "selectionMode",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		}
 	},
 	"types" : {
 		"Binding" : {
 			"id":"string",
 			"text": "string",
 			"type": "string"
 		},
 		"ContentControl" : {
 		 	"cannotDelete" : {"type": "boolean" },
 		 	"cannotEdit" : "boolean",
 		 	"color": "object",
 		 	"font": "object",
 		 	"placehoderText" : "string",
 		 	"removeWhenEdited" : "boolean",
 		 	"style" : "string",
 		 	"tag" : "string",
 			"text": "string",
 			"title" : "boolean"
 		}
 	}
}