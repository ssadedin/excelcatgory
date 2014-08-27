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

class GService {
    
    String applicationName

    /**
     * The scopes needed to browse and download google spreadsheets
     */
    final List scopes 
    
    /**
     * Create a google service helper object
     * 
     * @param applicationName
     * @param scopes
     */
    public GService(String applicationName, List scopes) {
        this.applicationName = applicationName
        this.scopes = scopes
    }
    
    static String secrets = '{"installed":{"auth_uri":"https://accounts.google.com/o/oauth2/auth","client_secret":"3kLxLZXUIE1WypjHGHJgojgw","token_uri":"https://accounts.google.com/o/oauth2/token","client_email":"","redirect_uris":["urn:ietf:wg:oauth:2.0:oob","oob"],"client_x509_cert_url":"","client_id":"831549207147.apps.googleusercontent.com","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs"}}'

    /** Directory to store user credentials. */
    final File DATA_STORE_DIR = new File(System.getProperty("user.home"), ".store/drive_sample");

    /**
     * Global instance of the {@link DataStoreFactory}. The best practice is to make it a single
     * globally shared instance across your application.
     */
    FileDataStoreFactory dataStoreFactory;

    /** Global instance of the JSON factory. */
    static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    /** Global instance of the HTTP transport. */
    static HttpTransport httpTransport;

    /** The actual Google Drive client */
    Drive client;

    /** Credentials used to authenticate */
    Credential credential

     /** Authorizes the installed application to access user's protected data. */
    Credential authorize() throws Exception {
        
        // load client secrets
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY,
                new ByteArrayInputStream(secrets.bytes).newReader());


        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, JSON_FACTORY, clientSecrets, scopes)
                .setDataStoreFactory(dataStoreFactory)
                .build();

        // authorize
        return new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
    }
    
    synchronized void initialize() {
        
        // Initialize the transport
        // This is shared globally (thread safe), so only initialize it if it is null
        if(httpTransport == null)
            httpTransport = GoogleNetHttpTransport.newTrustedTransport();

        // initialize the data store factory
        dataStoreFactory = new FileDataStoreFactory(DATA_STORE_DIR);

        // authorization
        Credential credential = authorize();

        // set up global Drive instance
        client = new Drive.Builder(httpTransport, JSON_FACTORY, credential).setApplicationName(applicationName).build();
    }

    /**
     * Query for spreadsheets containing the given title in their title
     * 
     * @param title
     * @return
     */
    List<GFile> query(String title, String type, String folderName) {
        
        GFile folder = this.find("title='$folderName' and mimeType='application/vnd.google-apps.folder'")
        if(!folder) 
            throw new RuntimeException("No folder $folderName could be found")
        
        DriveFileList request = client.files().list()
        request.q = "title contains '$title' and mimeType = '$type'"
          
        List<GFile> result = []

        // See https://developers.google.com/drive/v2/reference/files/list
        // for this logic
        while(true) {
            FileList files = request.execute()
            for(GFile f in files.getItems()) {
                result.add(f)
                // println "Found file $f.title with id $f.id"
            }        
            request.pageToken =  files.nextPageToken
            if(!request.pageToken)
                break
        }
        return result
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
    
    void eachDocument(String query, Closure c) {
        
          DriveFileList request = client.files().list()
          request.q = query
          
          // See https://developers.google.com/drive/v2/reference/files/list
          // for this logic
          while(true) {
              println "Executing query: $query"
              FileList files = request.execute()
              for(GFile f in files.getItems()) {
                  if(c(f)==Boolean.FALSE) {
                      return
                  }
              }        
              request.pageToken =  files.nextPageToken
              if(!request.pageToken)
                  break
          }
    }
    
    /**
     * Return an input stream from which to read the document with given ID
     * 
     * @param documentId    ID of document to return, as received from a GFile.id (see query)
     * @return InputStream
     */
    InputStream fetch(String documentId, String format) {
          client.requestFactory
                .buildGetRequest(new GenericUrl("https://docs.google.com/feeds/download/spreadsheets/Export?key=${documentId}&exportFormat=$format"))
                .execute()
                .content
    }
}
