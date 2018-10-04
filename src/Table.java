import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class Table {
    public static void main(String[] args) throws IOException
    {
        InputStream in=new FileInputStream("AYDS.docx");

        XWPFDocument docin=new XWPFDocument(in);
        List<XWPFTable> needs = docin.getTables();
        XWPFTable table = needs.get(1);
        table.getCTTbl();

    }
}
