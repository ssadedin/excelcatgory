import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.DataStoreFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.Drive.Files.List as DriveFileList;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.File as GFile
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.ListFeed
import com.google.gdata.data.spreadsheet.SpreadsheetEntry
import com.google.gdata.data.spreadsheet.SpreadsheetFeed
import com.google.gdata.data.spreadsheet.Worksheet;
import com.google.gdata.data.spreadsheet.WorksheetEntry

class GSheet extends GService {
    
    public GSheet(String applicationName) {
        super(applicationName, spreadSheetScopes)
    }
    
    /**
     * The scopes needed to browse and download google spreadsheets
     */
    static final List spreadSheetScopes =["https://spreadsheets.google.com/feeds", "https://docs.google.com/feeds"]
    
    public static void main(String[] args) {
        
        GSheet sheet = new GSheet("MG-JobPoller/1.0")
        sheet.initialize()
        
        List<GFile> files = sheet.query("Melbourne Genomics Job", "Pending")
        
        println "CSV for spreadsheet:\n\n" + sheet.fetch(files[0].id).text
        
        println "=" * 80
        
    }
    
    /**
     * Query for spreadsheets containing the given title in their title
     * 
     * @param title
     * @return
     */
    List<GFile> query(String title, String folderName) {
        query(title, 'application/vnd.google-apps.spreadsheet', folderName)
    }
    
    GFile find(String query, Closure c=null) {
        GFile result = null
        eachDocument(query) { GFile f ->
            result = f
            if(c == null || c(f))
                return false
        }
        return result
    }
    
    /**
     * Return an input stream from which to read the document with given ID
     * 
     * @param documentId    ID of document to return, as received from a GFile.id (see query)
     * @return InputStream
     */
    InputStream fetch(String documentId) {
       fetch(documentId, "csv")
    }
    
    List listNative() {
          // Let's try and read the spreadsheet
          SpreadsheetService service = new SpreadsheetService("MySpreadsheetService");
          service.setOAuth2Credentials(credential)

          URL sheetFeedUrl = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full")
          SpreadsheetFeed feed = service.getFeed(sheetFeedUrl,SpreadsheetFeed.class);

          List spreadsheets = feed.getEntries()

          for(SpreadsheetEntry s in spreadsheets){
              WorksheetEntry w = s.getDefaultWorksheet()

              println "Examining workseet " + w.getTitle().getPlainText() + w.listFeedUrl + " id = " + w.id
              URL listFeedUrl = w.getListFeedUrl();
              ListFeed listFeed = service.getFeed(listFeedUrl, ListFeed.class);

              listFeed.entries.each{row->
                  row.customElements.tags.each{ tag->
                      println "Column: ${tag} -- Value: ${row.getCustomElements().getValue(tag)}"
                  }
              }
              break
          }
        
    }
}
