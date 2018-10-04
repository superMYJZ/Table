import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class Table {
    public static void main(String[] args) throws IOException
    {
        InputStream in=new FileInputStream("AYDS.docx");

        XWPFDocument docin=new XWPFDocument(in);
        List<XWPFTable> needs = docin.getTables();
        XWPFTable table = needs.get(1);            //表
        // CTObject
        CTTblPr pr = table.getCTTbl().getTblPr();
        CTTblGrid grid = table.getCTTbl().getTblGrid();
        List<CTRow> trs = table.getCTTbl().getTrList();

        CTRow t1 = trs.get(0);
        List<CTTc> cs = t1.getTcList();
        CTTc c1 = cs.get(0);
        CTTcPr tcpr1 = c1.getTcPr();
        CTP x0 =c1.getPArray(0);
        CTR x1 = x0.getRArray(0);

        CTRPr x2 = x1.getRPr();
        CTFonts x3 = x2.getRFonts();//.getTList().get(0);
        System.out.println(x3.getHint());

        List<CTText> y2 = x1.getTList();
        String s1 = y2.get(0).getStringValue();
        // CTObject

        XWPFTableRow a = table.getRow(0);   //行
        a.setHeight(200);
        XWPFTableCell b = a.getCell(0);     //格
        b.setText("China");
        System.out.println(b.getText());
        CTTcPr x = b.getCTTc().getTcPr();
        System.out.println("A");
        docin.write(new FileOutputStream("aaa.docx"));
    }
}
