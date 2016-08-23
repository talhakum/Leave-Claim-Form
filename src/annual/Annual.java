package annual;


import com.microsoft.schemas.office.visio.x2012.main.CellType;
import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A dirty simple program that reads an Excel file.
 * @author www.codejava.net
 *
 */
public class Annual {
    public static CellStyle yillikFont(XSSFWorkbook workbook) {
        CellStyle styleNo=workbook.createCellStyle();
        styleNo.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderTop(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderRight(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderLeft(XSSFCellStyle.BORDER_THIN);
                return styleNo;
    }
    public static CellStyle noFont(XSSFWorkbook workbook) {
        CellStyle styleNo=workbook.createCellStyle();
        styleNo.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderTop(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderRight(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        styleNo.setAlignment(CellStyle.ALIGN_LEFT);
                XSSFFont font=workbook.createFont();
                font.setBold(true);
                styleNo.setFont(font);
                return styleNo;
    }
    public static CellStyle tarihFont(XSSFWorkbook workbook) {
        CellStyle styleNo=workbook.createCellStyle();
        styleNo.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderTop(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderRight(XSSFCellStyle.BORDER_THIN);
        styleNo.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        styleNo.setAlignment(CellStyle.ALIGN_LEFT);
        CreationHelper createHelper = workbook.getCreationHelper();
        styleNo.setDataFormat(
          createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
        styleNo.setAlignment(CellStyle.ALIGN_LEFT);
        return styleNo;
    }
    public static double izin(Date gelenTarih,String gun,String izinTur, String sheetName, String URL) throws FileNotFoundException, IOException, Exception {
      
        
            boolean test=false;
             boolean mzSelect=true;
            boolean select=true;
            double kalanİzin=999;
        
		//String excelFilePath = "C:\\Users\\talha\\Desktop\\deneme.xlsx";
    FileInputStream inputStream = new FileInputStream(new File(URL));
    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		Date tarih[]=new Date[150];
                String kullanilan[]=new String[150];
                String bakiye[]=new String[150];
                String aciklama[]=new String[150];
                int no[]=new int[150];
                String izinHakki="";
                double mazeretHakki=0;
                Date sonMazeret=null;
                Date SGKGiris=new Date();
		XSSFSheet firstSheet = workbook.getSheet(sheetName);
                Cell cell=null;
		Iterator<Row> iterator = firstSheet.iterator();
                int i=0;
            while (iterator.hasNext()) {
            // tarih[i]="";
            kullanilan[i]="";
            bakiye[i]="";
            aciklama[i]="";
            no[i]=0;       
                        Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                 cell = cellIterator.next();
                 if(cell.getColumnIndex()==0 && !cell.toString().isEmpty() && cell.getRowIndex()!=1) {
                      no[i]=(int)Double.parseDouble(cell.toString());
                     }
                if(cell.getColumnIndex()==1 && !cell.toString().isEmpty() && cell.getRowIndex()>1) {
                    tarih[i]=cell.getDateCellValue(); 
                }
                if(cell.getColumnIndex()==2 && !cell.toString().isEmpty()) {
                    kullanilan[i]=(cell.toString());
                }
                if(cell.getColumnIndex()==3 && !cell.toString().isEmpty()) {
                    bakiye[i]=(cell.toString());
                                }
                if(cell.getColumnIndex()==4 && !cell.toString().isEmpty()) {
                    aciklama[i]=(cell.toString());
                }
                if(cell.getRowIndex()==0 && cell.getColumnIndex()==4) {
                    SGKGiris=cell.getDateCellValue();
                }
            }
            if(!aciklama[i].isEmpty())
            i++;
        }
            for(int k=i-1;k>=0; k--) {
                System.out.println(no[k]+" "+tarih[k]+" "+kullanilan[k]+" "+bakiye[k]+" "+aciklama[k]);
                if((aciklama[k].equalsIgnoreCase("Yıllık İzin") || aciklama[k].equalsIgnoreCase("Yıllık İzin Hakkı") ||
                        aciklama[k].equalsIgnoreCase("Yıllık İzin Kullanımı")) && select ) {
                    select=false;
                 izinHakki=bakiye[k];
            }
                if(aciklama[k].startsWith("MZ") && mzSelect) {
                    mzSelect=false;
                    aciklama[k]=aciklama[k].substring(aciklama[k].indexOf("(") + 1);
                     aciklama[k] = aciklama[k].substring(0, aciklama[k].indexOf(")"));
                   if(aciklama[k].contains(","))
                        aciklama[k]=aciklama[k].replace(",", ".");
                     mazeretHakki=Double.parseDouble(aciklama[k]);
                     sonMazeret=tarih[k];
                }
            }
            if(sonMazeret==null) {
                mazeretHakki=0;
            }else {
            Calendar cal = Calendar.getInstance();    
            cal.set(Calendar.getInstance().get(Calendar.YEAR), SGKGiris.getMonth(), SGKGiris.getDate(),sonMazeret.getHours(),sonMazeret.getMinutes(),sonMazeret.getSeconds());
            SGKGiris=cal.getTime(); // date.setDate()'in hatalı olmasından bu yöntemi kullandım.
            
            cal.set(2000+gelenTarih.getYear()%100,sonMazeret.getMonth(),sonMazeret.getDate());
            sonMazeret=cal.getTime();
               
                
            if( (SGKGiris.compareTo(sonMazeret)>0) && (gelenTarih.compareTo(SGKGiris)>=0)){   //Son mazereti yıldönümünden önceyse mazeret hakkını 0'la.
                mazeretHakki=0;
            }
            }
            XSSFRow sheetrow = firstSheet.getRow(i);
            if(sheetrow == null){
             sheetrow = firstSheet.createRow(i);
            }
            cell = sheetrow.getCell(2);
        if(cell == null){
           cell = sheetrow.createCell(0);
           cell = sheetrow.createCell(1);
           cell = sheetrow.createCell(2);
           cell = sheetrow.createCell(3);
           cell = sheetrow.createCell(4);
        }
        switch(izinTur) {
            case "yillik": 
                if((Double.parseDouble(izinHakki))-(Double.parseDouble(gun))<0) {
                    throw new Exception("İZİN HAKKI YETERSİZ.\n Kalan İzin Hakkınız: "+ 
                            Math.floor(Double.parseDouble(izinHakki)*100)/100);
                }
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellValue("-"+gun);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(3);
                cell.setCellValue(String.valueOf( (Double.parseDouble(izinHakki))-(Double.parseDouble(gun)) ));
                kalanİzin=Math.floor(((Double.parseDouble(izinHakki))-(Double.parseDouble(gun)))*100)/100;
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Yıllık İzin Hakkı");
                cell.setCellStyle(yillikFont(workbook));
                break;
            case "mazeret":
                if(mazeretHakki==5 || ((mazeretHakki+Double.parseDouble(gun))>5)  ) { //aynı ay içerisinde mazeret kullanılamaz.
                    throw new Exception("MAZERET HAKKINIZ DOLMUŞTUR."); //Hata olarak yazdırılacak.
                }
                else if( ( (sonMazeret.getYear()%100==gelenTarih.getYear()%100) && sonMazeret.getMonth()==gelenTarih.getMonth() )){
                    if( (mazeretHakki+Double.parseDouble(gun))<=1  ){ //AYNI AY İÇERİSİNDE MAZERETİN 1'DEN AZ OLMA DURUMU.
                        cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-"+gun);
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("MZ("+String.valueOf(mazeretHakki+Double.parseDouble(gun))+")");
                kalanİzin=Math.floor((5-(mazeretHakki+Double.parseDouble(gun)))*100)/100;
                cell.setCellStyle(yillikFont(workbook));
                    }
                    else {
                        throw new Exception("AYNI AY İÇİNDE BİRDEN ÇOK MAZERET ALINAMAZ.\n Kalan Mazeret Hakkınız: "+Math.floor((5-mazeretHakki)*100)/100); //HATA OLARAK EKRANA YAZDIRILACAK.!
                    }
            }
                else {
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1); 
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-"+gun);
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("MZ("+String.valueOf(mazeretHakki+Double.parseDouble(gun))+")");
                kalanİzin=Math.floor((5-(mazeretHakki+Double.parseDouble(gun)))*100)/100;
                cell.setCellStyle(yillikFont(workbook));
                }
                break;
            case "rapor":
                 cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-"+gun);
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Rapor");
                cell.setCellStyle(yillikFont(workbook));
                break;
            case "dogumizni":
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-112");
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Doğum İzni");
                cell.setCellStyle(yillikFont(workbook));
                break;
            case "evlilik":
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-3");
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Kanuni Evlilik İzni");
                cell.setCellStyle(yillikFont(workbook));
                break;
            case "babalik":
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-5");
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Babalık İzni");
                cell.setCellStyle(yillikFont(workbook));
                break;
            case "sutizni":
                cell=sheetrow.getCell(0);
                cell.setCellValue(no[i-1]+1);
                cell.setCellStyle(noFont(workbook));
                cell=sheetrow.getCell(1);
                cell.setCellValue(gelenTarih);
                cell.setCellStyle(tarihFont(workbook));
                cell=sheetrow.getCell(2);
                cell.setCellStyle(yillikFont(workbook));
                cell.setCellValue("-"+gun);
                cell=sheetrow.getCell(3);
                cell.setCellStyle(yillikFont(workbook));
                cell=sheetrow.getCell(4);
                cell.setCellValue("Süt İzni");
                cell.setCellStyle(yillikFont(workbook));
                break;
        }
            inputStream.close();
            FileOutputStream outFile =new FileOutputStream(new File("C:\\Users\\talha\\Desktop\\deneme.xlsx"));
            workbook.write(outFile);
            outFile.close();
            return kalanİzin;
    } 
	public static void main(String[] args) throws IOException, IOException  {
            Date deneme=new Date();
                Calendar calendar = Calendar.getInstance();    
            calendar.set(2017, 03, 05);
            deneme=calendar.getTime(); // Parametre ile alacağımız tarih buraya gelecek. !
            System.out.println(deneme);
    }
}