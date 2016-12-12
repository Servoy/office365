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
		"insertContentControl" : {
					"parameters": [{
					"name": "contentControlOptions",
					"type": "ContentControl"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"getAllContentControls" : {
			    "parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "object[]"
		},
		"goToContentControl" : {
			    "parameters": [{
					"name": "title",
					"type": "string"
				},{
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
		"setContentControlText": {
			"parameters": [{
					"name": "title",
					"type": "string"
				},{
					"name": "text",
					"type": "string"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "int"
		},
		"deleteContentControl": {
			"parameters": [{
					"name": "title",
					"type": "string"
				},{
					"name": "keepContent",
					"type": "boolean"
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
		"getSearchResults": {
			"returns" : "ContentControl[]"
		},
		"selectSearchResult" : {
			"parameters": [{
					"name": "index",
					"type": "int"
				},{
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
 		 	"cannotDelete" : "boolean",
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