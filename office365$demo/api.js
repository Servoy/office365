/**
 * @type {String}
 *
 *
 * @properties={typeid:35,uuid:"4D438161-55AD-46CC-8353-04D300FE55F5"}
 */
var token = "EwAAA61DBAAUGCCXc8wU/zFu9QnLdZXy+YnElFkAAUeLABS3wh9solWuWl53d4JpDu+1ICrv2S4TkiPVAVNYk1+K19tHYkb8+Bz9mHoBLNdhksvd829Yhc9Qjwt71k3tW8USjGNNbZXTAJm4d6DaHhrkuPhgMPxpvl9t3JtEgey+c0JOwhaWs8UO0O+k6uTzVH8E3bh+JQiMiTjVZJMj3f5ltzN7LHTvKcywHKZ7nRivbUtZICTdYf+WrwdiF+y9YBpsdQzqLpKE5/wFM3H3oy97pd8cXiYFhSJfFqJnjr/eiQigqkZ2wUIrhCSuVWj34QpnrgZ/YFrdrtTWw8swuv4U4LZLbxoQZhj0R4fApij9OcPXHMPCTkdF4/tZZZkDZgAACDKS5C5ltM4B0AFIQXHHjB0pWkFjmNKOKUFysOzVYfcaoV1u96aXMDlEctmEi8B7eVl1RRDbY8vtcjKQ+TBy05L50/PvLdwJ81R0VvGZQo1/n3rGpkJaWzHYGECwu3qavK+yck7gPoTbg5AAPMyzkk4gA/fbLrixjxi0eb+4Hprqv5qqEqpX7ypQFoJcU4CfMi4ExGjrtASYasN6NVfAwqE69LZ3M8WBmv/f7kpLfzBn2RhB0xqYkruDdDXU5FAMYoEeW7xw+QAf0RwYmiaLp+dp2/A1Hdku5LAL9rKBoWdX7Ert6lsb1FgX58kR/+B4WVe2Jlpj89xdfsaTbkQM7aFm9dIHqdrX/u/wOSj79OyGgB9DlF1ZpopJemuu8MBa4xAPo/JMsv4Z8nEumnqe8PlYv3ZuDEYj7hjhtO03PPEpRWABX6yTRj1r32rXskRItb8mybH978aLke7USqXIGsaWhCouBzxh1PVCFuNRcCRVnpoabXxaOzs1Veu224qIizKT195ZHJ+aA10ExMm9uslInAw5x4CqYXAwcwg9qKDImT/3dh65BNX8y89rA9QEzMDcFhYKs1hjF1pH8RD1YjioDFRXxPWVk44RjMvi+rRGvGsZ8xrcHFKTyPYB"
	
/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"935C2053-B0AD-4838-83BA-41814F05292E"}
 */
function onAction(event) {
	// TODO Auto-generated method stub
	getDrive();
	var file = plugins.file.showFileOpenDialog(oneDriveUpload)

}


/**
 * TODO generated, please specify type and doc for the params
 * @param file
 * @param filename
 * @param file_content
 * @properties={typeid:24,uuid:"E42F1031-537D-44D9-A208-20BFA6F946BB"}
 */
function oneDriveUpload(file,filename,file_content){
	
	application.output(file)
	application.output(filename)
	application.output(file_content)
	var $url   = "https://graph.microsoft.com/v1.0/";
	var $client   = plugins.http.createNewHttpClient();
	var $access_token   = RefreshToken(_to_app_user$current_user.office_365_refresh_token);

	var post_url = '/me/drive/items/'+parent_id+':/'+filename+':/content';

	var $post   = $client.createPutRequest($url + encodeURI(post_url));
	$post.addHeader("Authorization", "Bearer " + token)
	$post.addHeader("Content-Type", "text/plain");
	//   $post.addHeader("Content-Type", "application/json");
	//   $post.addHeader("Content-Type", "application/octet-stream");
	//$post.setBodyContent(file_content)
	$post.setFile(file)
	
	var executeRequest   = $post.executeRequest();
	return {
	'getStatusCode': executeRequest.getStatusCode(),
	'getResponseBody': JSON.parse(executeRequest.getResponseBody())
	};


	//   /me/drive/items/{parent-id}:/{filename}:/content
}


/**
 * @properties={typeid:24,uuid:"849D1AE0-F87D-4F7E-98CC-E7DCE291CAD3"}
 */
function getDrive() {

	var url   = "https://api.onedrive.com/v1.0/drive";
	var $client   = plugins.http.createNewHttpClient();
	//var $access_token   = RefreshToken(_to_app_user$current_user.office_365_refresh_token);

	//var post_url = '/me/drive/items/'+parent_id+':/'+filename+':/content';

	var $post   = $client.createGetRequest(url)//($url + encodeURI(post_url));
	$post.addHeader("Authorization", "Bearer " + token)
	//$post.addHeader("Content-Type", "text/plain");
	var reg = $post.executeRequest()
	application.output(reg.getStatusCode());
	application.output(reg.getResponseBody());
}
