import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.xwpf.usermodel.BodyElementType.PARAGRAPH;

public class Table
{
	public static void main(String[] args) throws IOException, XmlException
	{
		InputStream in = new FileInputStream("AYDS.docx");

		XWPFDocument docin = new XWPFDocument(in);
		List<XWPFTable> tables = docin.getTables();
		List<CTRow> trs;
		List<CTTc> tcs;
		CTTcPr tcpr;
		CTRow tr;
		for (int j = 0; j < tables.size(); j++)             //遍历每个Table
		{
			if (j==0)continue;                              //第一个Table不做处理
			trs=tables.get(j).getCTTbl().getTrList();
			for (int i = 0; i < trs.size(); i++)            //遍历Table的行
			{
				tr = trs.get(i);
				tcs = tr.getTcList();
				if (i == 0)                                 //第一行只删除左右
				{
					for (CTTc tc : tcs)
					{
						tcpr = tc.getTcPr();
						CTTcBorders tcb = tcpr.addNewTcBorders();
						CTBorder left = tcb.addNewLeft();
						left.setVal(STBorder.Enum.forInt(1));
						CTBorder right = tcb.addNewRight();
						right.setVal(STBorder.Enum.forInt(1));
					}
				}
				else if (i == (trs.size() - 1))             //最后一行只删除上左右
				{
					for (CTTc tc : tcs)
					{
						tcpr = tc.getTcPr();
						CTTcBorders tcb = tcpr.addNewTcBorders();
						CTBorder top = tcb.addNewTop();
						top.setVal(STBorder.Enum.forInt(1));
						CTBorder left = tcb.addNewLeft();
						left.setVal(STBorder.Enum.forInt(1));
						CTBorder right = tcb.addNewRight();
						right.setVal(STBorder.Enum.forInt(1));
					}
				}
				else for (CTTc tc : tcs)                    //其他全部处理
					{
						tcpr = tc.getTcPr();
						CTTcBorders tcb = tcpr.addNewTcBorders();
						CTBorder top = tcb.addNewTop();
						top.setVal(STBorder.Enum.forInt(1));
						CTBorder bottom = tcb.addNewBottom();
						bottom.setVal(STBorder.Enum.forInt(1));
						CTBorder left = tcb.addNewLeft();
						left.setVal(STBorder.Enum.forInt(1));
						CTBorder right = tcb.addNewRight();
						right.setVal(STBorder.Enum.forInt(1));
					}

			}
		}
		docin.write(new FileOutputStream("aaa.docx"));
		docin.close();
	}
}
