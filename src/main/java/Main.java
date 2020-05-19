import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException, DocumentException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream("./src/main/resources/list.xlsx"));
        BaseFont bf = BaseFont.createFont("./src/main/resources/times.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        BaseFont bf1 = BaseFont.createFont("./src/main/resources/timesi.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        Sheet sheet = workbook.getSheet("Лист1");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            PdfReader pdf = new PdfReader("C:\\Users\\xopow\\OneDrive\\Desktop\\sertificate.pdf");

            String name = sheet.getRow(i).getCell(1).getStringCellValue();
            String theme = sheet.getRow(i).getCell(2).getStringCellValue();

            String filename = null;

            if (theme.length() > 50) {
                filename = (name + "_" + theme.substring(0, 50))
                        .replaceAll("[\"?:/*<>|]", "");
            }
            else {
                filename = (name + "_" + theme)
                        .replaceAll("[\"?:/*<>|]", "");
            }

            System.out.println(filename + " started");

            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\xopow\\OneDrive\\Desktop\\result\\" + filename + ".pdf");

            PdfStamper stamper = new PdfStamper(pdf, fileOutputStream);

            PdfContentByte contentByte = stamper.getOverContent(1);

            contentByte.beginText();
            contentByte.setFontAndSize(bf, 16);
            contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, name, 424, 231, 0);

            if (theme.length() < 130) {
                contentByte.setFontAndSize(bf1, 12);
                contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, theme, 424, 214, 0);
                contentByte.endText();
            }
            else {
                System.out.println("1");
                int spaceIndex = 0;
                for (int j = 129; theme.charAt(j) != ' '; j--) {
                    if (theme.charAt(j - 1) == ' ') {
                        spaceIndex = j - 1;
                        break;
                    }
                }
                contentByte.setFontAndSize(bf1, 10);
                contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, theme.substring(0, spaceIndex), 424, 220, 0);
                contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, theme.substring(spaceIndex), 424, 210, 0);
            }

            stamper.close();
            fileOutputStream.close();

            System.out.println(filename + " completed");
        }
    }
}