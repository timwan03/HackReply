
function dummyCallback(asyncResult:Office.AsyncResult)
{
    console.log(JSON.stringify(asyncResult));
}

function ewsInstantSend(id:string, changekey:string, body:string, fReplyAll:boolean):string
{
	let type:string;
	if (fReplyAll)
		type = "ReplyAllToItem";
	else
		type = "ReplyToItem";
	
	let result:string = 
	'<?xml version="1.0"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
		'<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
		'<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
		'</soap:Header>' +
		'<soap:Body>' +
	      '<m:CreateItem MessageDisposition="SendAndSaveCopy">' +
	        '<m:Items>' +
	          '<t:' + type + '>' +
	            '<t:ReferenceItemId Id="' + id + '" ChangeKey="' + changekey + '"/>' +
	            '<t:NewBodyContent BodyType="HTML">' + body + '</t:NewBodyContent>' +
	          '</t:' + type + '>' +
	        '</m:Items>' +
	      '</m:CreateItem>' +
	    '</soap:Body>' +
	  '</soap:Envelope>';

	return result;
}

export function onClickInstantSend(itemId:string, bodyText:string, fReplyAll:boolean) 
{
	var getItemRequestString;

	getItemRequestString = getItemRequest(itemId);

	Office.context.mailbox.makeEwsRequestAsync(getItemRequestString, getItemCallback, {"body":bodyText, "itemId":itemId, "fReplyAll":fReplyAll});
}

function getItemCallback(obj:Office.AsyncResult)
{
	let changeKey:string = getChangeKey(obj.value);
    
    let bodyString:string = '<font face = "Calibri">' + obj.asyncContext.body + '</font>';
    bodyString = bodyString.replace(/</g,"&lt;").replace(/>/g,"&gt;");

    let ewsString:string = ewsInstantSend(obj.asyncContext.itemId, changeKey, bodyString, obj.asyncContext.fReplyAll);
    Office.context.mailbox.makeEwsRequestAsync(ewsString, dummyCallback);

    return false;
}

function getChangeKey(r:string) 
{
    var i = r.indexOf('ChangeKey="');
    if (i > 0) {
        var changeKey = r.substring(i + 11);
        i = changeKey.indexOf('"');
        if (i > 0) {
            return changeKey.substring(0, i);
        }
    }
    return null;
}

function getItemRequest(id:string) : string 
{
	let result:string = 
		'<?xml version="1.0"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
		'<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
			'<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
		'</soap:Header>' +
			'  <soap:Body>' +
			'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
			'      <ItemShape>' +
			'        <t:BaseShape>IdOnly</t:BaseShape>' +
			'        <t:AdditionalProperties>' +
			'            <t:FieldURI FieldURI="item:Subject"/>' +
			'        </t:AdditionalProperties>' +
			'      </ItemShape>' +
			'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
			'    </GetItem>' +
			'  </soap:Body>' +
			'</soap:Envelope>';

	return result;
}