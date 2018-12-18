{
	"name": "office365-outlook",
	"displayName": "outlook",
	"version": 1,
	"definition": "office365/outlook/outlook.js",
	"libraries": 
	[
		
	],

	"model": 
	{
		"ATTACHMENT_TYPE": 
		{
			"type": "attachmentType",
			"default": 
			{
				
			}
		},

		"ATTACHMENT_CONTENT_FORMAT": 
		{
			"type": "AttachmentContentFormat",
			"default": 
			{
				
			}
		},

		"COERCION_TYPE": 
		{
			"type": "coercionType",
			"default": 
			{
				
			}
		},

		"searchResult": 
		{
			"type": "object",
			"pushToServer": "allow",
			"tags": 
			{
				"scope": "private"
			}
		}
	},

	"api": 
	{
		"helloworld": 
		{
			"parameters": 
			[
				{
					"name": "text",
					"type": "string"
				}
			]
		},

		"addRecipients": 
		{
			"parameters": 
			[
				{
					"name": "recipients",
					"type": "recipient[]"
				},

				{
					"name": "onError",
					"type": "function"
				}
			]
		},

		"addCCRecipients": 
		{
			"parameters": 
			[
				{
					"name": "recipients",
					"type": "recipient[]"
				},

				{
					"name": "onError",
					"type": "function"
				}
			]
		},

		"addBCCRecipients": 
		{
			"parameters": 
			[
				{
					"name": "recipients",
					"type": "recipient[]"
				},

				{
					"name": "onError",
					"type": "function"
				}
			]
		},

		"getSelectedMessageBodyType": 
		{
			"parameters": 
			[
				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "object"
		},

		"getSelectedMessageBody": 
		{
			"parameters": 
			[
				{
					"name": "coercionType",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "object"
		},

		"displayNewMessageForm": 
		{
			"parameters": 
			[
				{
					"name": "toRecipients",
					"type": "string[]"
				},

				{
					"name": "ccRecipients",
					"type": "string[]"
				},

				{
					"name": "subject",
					"type": "string"
				},

				{
					"name": "htmlBody",
					"type": "string"
				},

				{
					"name": "attachments",
					"type": "attachment[]"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "object"
		},

		"getSelectedMessageSubject": 
		{
			"parameters": 
			[
				{
					"name": "onError",
					"type": "function"
				}
			]
		},

		"setSubject": 
		{
			"parameters": 
			[
				{
					"name": "subject",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			]
		},

		"setBodyText": 
		{
			"parameters": 
			[
				{
					"name": "body",
					"type": "string"
				},

				{
					"name": "coercionType",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "boolean"
		},

		"prependBodyText": 
		{
			"parameters": 
			[
				{
					"name": "body",
					"type": "string"
				},

				{
					"name": "coercionType",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "boolean"
		},

		"setSelectedText": 
		{
			"parameters": 
			[
				{
					"name": "body",
					"type": "string"
				},

				{
					"name": "coercionType",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "boolean"
		},

		"addAttachment": 
		{
			"parameters": 
			[
				{
					"name": "file",
					"type": "string"
				},

				{
					"name": "attachmentName",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "boolean"
		},

		"getAttachments": 
		{
			"parameters": 
			[
				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "object[]"
		},

		"getAttachmentContent": 
		{
			"parameters": 
			[
				{
					"name": "attachmentId",
					"type": "string"
				},

				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "object"
		},

		"saveItem": 
		{
			"parameters": 
			[
				{
					"name": "onError",
					"type": "function"
				}
			],

			"returns": "string"
		},

		"getEWSServer": 
		{
			"parameters":
			[
				{
					"name": "onError",
					"type": "function"
				}
			],
			
			"returns": "string"
		}
	},

	"types": 
	{
		"coercionType": 
		{
			"TEXT": 
			{
				"type": "string",
				"default": "text"
			},

			"HTML": 
			{
				"type": "string",
				"default": "html"
			}
		},

		"attachmentType": 
		{
			"FILE": 
			{
				"type": "string",
				"default": "file"
			},

			"ITEM": 
			{
				"type": "string",
				"default": "item"
			},

			"CLOUD": 
			{
				"type": "string",
				"default": "cloud"
			}
		},

		"AttachmentContentFormat": 
		{
			"BASE64": 
			{
				"type": "string",
				"default": "base64"
			},

			"EML": 
			{
				"type": "string",
				"default": "eml"
			},

			"ICALENDA": 
			{
				"type": "string",
				"default": "iCalendar"
			},

			"URL": 
			{
				"type": "string",
				"default": "url"
			}
		},

		"attachment": 
		{
			"attachmentType": 
			{
				"type": "attachmentType"
			},

			"contentType": 
			{
				"type": "string"
			},

			"id": 
			{
				"type": "string"
			},

			"isInline": 
			{
				"type": "boolean"
			},

			"name": 
			{
				"type": "string"
			},

			"size": 
			{
				"type": "double"
			}
		},

		"attachmentContent": 
		{
			"content": 
			{
				"type": "string"
			},

			"format": 
			{
				"type": "AttachmentContentFormat"
			}
		},

		"recipient": 
		{
			"displayName": 
			{
				"type": "string"
			},

			"emailAddress": 
			{
				"type": "string"
			},

			"recipientType": 
			{
				"type": "string"
			}
		}
	}
}