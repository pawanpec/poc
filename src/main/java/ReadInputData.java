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

	public static JSONArray readInputDataFromAPI(String apiUrl) {
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