{
	"name": "office365-word",
	"displayName": "word",
	"version": 1,
 	"definition": "office365/word/word.js",
	"libraries": [{
	
	}],
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
			"return" : "string"
		},
		"getDocumentUrl": {
	    	"parameters": [{
					"name": "onError",
					"type": "function",
					"optional" : true
				}],
			"return" : "string"
		},
		"getSelectedData": {
	    	"parameters": [{
					"name": "onSuccess",
					"type": "function"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}]
		},
		"setSelectedData": {
	    	"parameters": [{
					"name": "text",
					"type": "object"
				},{
					"name": "type",
					"type": "string"
				},{
					"name": "onSuccess",
					"type": "function"
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}]
		},
		"insertText": {
	    	"parameters": [{
					"name": "text",
					"type": "string"
				},{
					"name": "onSuccess",
					"type": "function",
					"optional" : true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}]
		},
		"insertMergeField": {
	    	"parameters": [{
					"name": "text",
					"type": "string"
				},{
					"name": "onSuccess",
					"type": "function",
					"optional" : true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}]
		},
		"writeDocx": {
	    	"parameters": [{
					"name": "xml",
					"type": "string"
				},{
					"name": "onSuccess",
					"type": "function",
					"optional" : true
				},{
					"name": "onError",
					"type": "function",
					"optional" : true
				}]
		}
 	}
}