{
	"name": "office365-word",
	"displayName": "word",
	"version": 1,
 	"definition": "office365/word/word.js",
	"libraries": [
	],
	"model":
	{
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
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
				
		},
		"setSelectedData": {
	    	"parameters": [{
					"name": "text",
					"type": "object"
				},{
					"name": "type",
					"type": "string"
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
		"insertBindingToBody" : {
			"parameters": [{
					"name": "text",
					"type": "string"
				},{
					"name": "id",
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
			"returns" : "string"
		},
		"insertContentControl" : {
					"parameters": [{
					"name": "tag",
					"type": "string"
				},{
					"name": "contentControlOptions",
					"type": "object",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
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
		},
		"searchBody": {
	    	"parameters": [{
					"name": "searchText",
					"type": "string"
				},{
					"name": "searchOptions",
					"type": "object",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "object"
		},
		"getAllBindings" : {
			    "parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "Binding[]"
		},
		"getBinding" : {
			    "parameters": [{
					"name": "id",
					"type": "string"
					},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "Binding"
		},
		"addBindingFromSelection" : {
			    "parameters": [{
					"name": "bindingType",
					"type": "string"
				},{
					"name": "id",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"addBindingFromPrompt" : {
			    "parameters": [{
					"name": "bindingType",
					"type": "string"
				},{
					"name": "id",
					"type": "string",
					"optional": true
				},{
					"name": "prompText",
					"type": "string",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"goToBinding" : {
			    "parameters": [{
					"name": "id",
					"type": "string"
				},{
					"name": "selectBinding",
					"type": "boolean",
					"optional": true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"setBindingData": {
	    	"parameters": [{
					"name": "id",
					"type": "string"
				},{
					"name": "text",
					"type": "string"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "boolean"
		},
		"setTagText": {
			"parameters": [{
					"name": "tag",
					"type": "string"
				},{
					"name": "text",
					"type": "string"
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
 		}
 	}
}