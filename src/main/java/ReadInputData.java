import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import org.json.simple.JSONObject;
import org.json.simple.JSONArray;
import org.json.simple.parser.JSONParser;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;


public class ReadInputData {
	 public static JSONParser parser = new JSONParser();
	public static void main(String[] args) {
	        getInputData();
	    }
	@SuppressWarnings("finally")
	public static JSONArray getInputData() {
		JSONObject jsonObject=null;
		JSONArray jsonarray=null;
		try {
			/*jsonarray = (JSONArray) parser.parse(new FileReader("input.json"));
			System.out.println(">>>>>>>>>>>>" + jsonarray);
			*/
			 // To connect to mongodb server
	         MongoClient mongoClient = new MongoClient( "localhost" , 27017 );
				
	         // Now connect to your databases
	         DB db = mongoClient.getDB( "TVLArticlesDB" );
	         System.out.println("Connect to database successfully" + db);
	        
	         DBCollection coll = db.getCollection("articles_3_2016");
	         
	         DBObject query = new BasicDBObject();
	         query.put("tags.0.ISIN", "US30303M1027");
	         BasicDBObject fields = new BasicDBObject("tvl2Cats",1).append("articlePubDateMs", 1);
	         DBObject orderBy=new BasicDBObject("articlePubDateMs",1);
			 List<DBObject> arr = (ArrayList<DBObject>) coll.find(query,fields).sort(orderBy).toArray();

	         Object object=null;
	         JSONParser jsonParser=new JSONParser();
	         object=jsonParser.parse(arr.toString());
	         jsonarray=(JSONArray) object;
	         System.out.println(">>>>>>>>>>>>>. abc :::"+arr.size() +":::"+jsonarray);	

		} catch (Exception e) {
		    e.printStackTrace();
		} finally {
			return jsonarray;	
		}
	}

	}