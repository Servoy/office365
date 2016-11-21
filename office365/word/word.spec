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
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"insertMergeField": {
	    	"parameters": [{
					"name": "text",
					"type": "string"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		},
		"writeDocx": {
	    	"parameters": [{
					"name": "xml",
					"type": "string"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"returns" : "string"
		}
 	}
}