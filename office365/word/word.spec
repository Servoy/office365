{
	"name": "office365-word",
	"displayName": "word",
	"version": 1,
 	"definition": "office365/word/word.js",
	"libraries": [],
	"model":
	{
 	},
 	"api":
 	{
 	 	"isOfficeEnabled": {
	    	"parameters": [],
			"returns" : "boolean"
		},
 		"getText": {
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
		 "getHtml": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		 "getOoxml": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
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
		"insertText": {
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
		"insertOoxml": {
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
		"select": {
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
		"search": {
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
		"gotoBinding" : {
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