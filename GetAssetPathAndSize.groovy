import com.day.cq.dam.api.Asset
import javax.jcr.Node
import com.day.cq.dam.api.AssetManager;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.text.SimpleDateFormat;
import org.apache.commons.io.IOUtils;
import java.text.DecimalFormat;


def paths = ["/content/dam"]

Date date = new Date();
SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy'T'HH-mm-ss");
def dateTimeString = dateFormat.format(date) +"/";
def location = "/content/dam/AssetPathAndSize/"+dateTimeString


HSSFWorkbook workbook = new HSSFWorkbook();
sheet = workbook.createSheet("AssetsPathAndSize");	
assetManager = resourceResolver.adaptTo(AssetManager)
HSSFRow rowhead = sheet.createRow(0);
    def columnCount  =0;
    Cell assetPathCell = rowhead.createCell(columnCount++);
    assetPathCell.setCellValue("Asset Path");
  
    Cell assetSizeCell = rowhead.createCell(columnCount++);
    assetSizeCell.setCellValue("Size");
  	

HSSFRow row;
def rowCount = 1;
def totalSize =0;
for(i in paths) {
    def predicates = [path:i, type: "dam:Asset"]
    def query = createQuery(predicates)
    query.hitsPerPage = 500
    def result = query.result
    
    result.hits.each { hit ->
        def path=hit.node.path

        Resource res = resourceResolver.getResource(path)
        Resource jcrResource = resourceResolver.getResource(path+"/jcr:content/metadata")
        if(res!=null) {
          Asset asset = res.adaptTo(Asset);
            Node jcrNode = jcrResource.adaptTo(javax.jcr.Node)
            if(asset != null &&  asset.getMetadataValue("dam:size") != null) {
            row = sheet.createRow(rowCount++);
            def innerColumnCount  = 0;
                Cell assetPathValueCell = row.createCell(innerColumnCount++);
                assetPathValueCell.setCellValue(path);
                 Cell assetSizeValueCell = row.createCell(innerColumnCount++);
                 def size = Integer.valueOf(asset.getMetadataValue("dam:size"));
                 totalSize = totalSize + size
                assetSizeValueCell.setCellValue(readableFileSize(size));
               
           }
        }
      
    }
 
}
def totalReadableSize =  readableFileSize(totalSize)
row = sheet.createRow(rowCount++);
Cell assetTotalValueCell = row.createCell(0);
assetTotalValueCell.setCellValue("Total Value of Asset Size :" +totalReadableSize);
createDamFileForExcel(workbook, location);
 
// method to create spreadsheet asset in dam
def createDamFileForExcel(def workbook, def location) {
    def filename = "AssetMetadata.xls";
    def baos = new ByteArrayOutputStream();
    workbook.write(baos);
    def bais = new ByteArrayInputStream(baos.toByteArray());
    def asset = assetManager.createAsset(location + filename, bais, "application/vnd.ms-excel", true);
    bais.close();
    baos.close();
	println location + filename
}


def readableFileSize(size) {
		if (size <= 0L)
			return "0";
		def units = [ "B", "kB", "MB", "GB", "TB" ];
		int digitGroups = (int) (Math.log10(size) / Math.log10(1024.0D));
	return (new DecimalFormat("#,##0.#")).format(size / Math.pow(1024.0D, digitGroups)) + " " + units[digitGroups];
	}
