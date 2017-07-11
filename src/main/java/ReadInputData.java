import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;

public class ReadInputData {
	public static JSONParser parser = new JSONParser();

	public static void main(String[] args) {
		//getInputData();
		readInputDataFromAPI();
	}

	@SuppressWarnings("finally")
	public static JSONArray getInputData() {
		JSONObject jsonObject = null;
		JSONArray jsonarray = null;
		try {
			/*
			 * jsonarray = (JSONArray) parser.parse(new
			 * FileReader("input.json")); System.out.println(">>>>>>>>>>>>" +
			 * jsonarray);
			 */
			// To connect to mongodb server
			MongoClient mongoClient = new MongoClient("localhost", 27017);

			// Now connect to your databases
			DB db = mongoClient.getDB("TVLArticlesDB");
			System.out.println("Connect to database successfully" + db);

			DBCollection coll = db.getCollection("articles_3_2016");

			DBObject query = new BasicDBObject();
			query.put("tags.0.ISIN", "US30303M1027");
			BasicDBObject fields = new BasicDBObject("tvl2Cats", 1).append(
					"articlePubDateMs", 1);
			DBObject orderBy = new BasicDBObject("articlePubDateMs", 1);
			List<DBObject> arr = (ArrayList<DBObject>) coll.find(query, fields)
					.sort(orderBy).toArray();

			Object object = null;
			JSONParser jsonParser = new JSONParser();
			object = jsonParser.parse(arr.toString());
			jsonarray = (JSONArray) object;
			System.out.println(">>>>>>>>>>>>>. abc :::" + arr.size() + ":::"
					+ jsonarray);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			return jsonarray;
		}
	}

	public static JSONArray readInputDataFromAPI() {
		String apiUrl = "https://devapi.insight360.io/v3/data/articles?start_date=2016-04-01&end_date=2016-04-30&ISIN=US30303M1027&token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczovL2RldmF1dGguaW5zaWdodDM2MC5pbyIsInN1YiI6InNzZGUiLCJleHAiOiIyMDE4LTA3LTA2VDIwOjQ2OjM0LjQzNVoiLCJpYXQiOiIyMDE3LTA3LTA2VDIwOjQ2OjM0LjQzNVoiLCJuYW1lIjoiWW9naURRIiwiZW5jIjoiZTA0ZmIxOWQ0MGQ3OTkzM2QwYjM3ZDZmYzRhM2MwMzdkNTQwZWZjMzk4Y2U2OWNlZjc2YTM5NjRkMzdhYzBiNWJmMjA2NzAxNzg2MzcyODZkZjk0ZDlhMzg4MzNmYWJlNWE5ODUzMjI3M2Q3NmM5YTgxMmJkOGM4NjQyZmQ2YmIwMGVmOTAzZTM0OTI2YTMzNTE0YWY0YmI1ZjVlOTM2MGJiOGZhYjYyOWI5ODgxMjM4ZjBhMDEzZjAyZWE1MGQwZDZlNzUxNDZiZjhjYjZhOTRkZWMyMTk4MjBlOGQ0ZWE3NjY1NjZlOTViMTRmN2YxZWVkOWFlNWNkYTJjNDJiOWI1MzU4YzE1OTgxM2Y1Y2E5ZmZiMzlkYWFkNDUyNDg4NmY5Mzc2ZTY0NjQ3YTUxZTA5OTIzNGJjZGE2NWI2N2NjMDNmODRkZDU5ZTNhMzNkZmVmMDhmMDBkNjEwZDc0ZGJkMTNhNzA2ZDhkYTgzNWEwZjJiOWZlMjY3MGY2MGYxZjFmNmM3MDRjNDBlMzI4YzA5MWFlYmFjN2NkY2MxN2M0MGRkMTZkYmJmYTYzMDhmYzMyMzRiIn0.J-lgV5EsDCaJlZibkvgr1vgwvAC0yhitB_yN7SrkbJ4";
		String apiResponse = CommonUtility.getAPIRespose(apiUrl);

		JSONArray jsonArr = null;
		JSONArray newJsonArray = new JSONArray();
		DateFormat format = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
		try {
			JSONParser jsonParser = new JSONParser();
			jsonArr = (JSONArray) jsonParser.parse(apiResponse);
			
			for (int i = 0; i < jsonArr.size(); i++) {
				JSONObject newObj = new JSONObject();
				 JSONObject jsonObj = (JSONObject) jsonArr.get(i);
				 newObj.put("tvl2Cats", jsonObj.get("tvl2Cats"));
				 Date date = format.parse((String)jsonObj.get("pubDate"));
				 newObj.put("articlePubDateMs",date.getTime());
				 newJsonArray.add(newObj);
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}		
		return newJsonArray;
	}

}