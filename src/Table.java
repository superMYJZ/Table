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
	/*测试代码********************************************************************************************
	public static void main(String[] args) throws IOException, XmlException
	{
		InputStream in = new FileInputStream("aaa.docx");
		XWPFDocument docin = new XWPFDocument(in);
		toThreeWireTable(docin);
		docin.write(new FileOutputStream("bbb.docx"));
		docin.close();
	}
	测试代码*********************************************************************************************/

	//传入XWPFDocument将文档内的所有Table转化为三线表
	public static void toThreeWireTable(XWPFDocument docin)
	{

		List<XWPFTable> tables = docin.getTables();
		List<CTRow> trs;
		List<CTTc> tcs;
		CTTcPr tcpr;
		CTRow tr;
		for (int j = 0; j < tables.size(); j++)             //遍历每个Table
		{
			if (j==0)continue;                              //第一个Table不做处理 是封面的Table
			trs=tables.get(j).getCTTbl().getTrList();       //获取表的行

			for (int i = 0; i < trs.size(); i++)            //遍历Table的行
			{
				tr = trs.get(i);
				tcs = tr.getTcList();
				if (i == 0)                                 //第一行只删除左右
				{
					for (CTTc tc : tcs)
					{
						tcpr = tc.getTcPr();
						CTTcBorders stcb = tcpr.getTcBorders();                //获取到原修饰体

						if (stcb!=null)         //判断是否有
						{
							if(stcb.getLeft()!=null){       //判断是否有左
								stcb.getLeft().setVal(STBorder.NIL);
							}else
							{
								CTBorder left = stcb.addNewLeft();
								left.setVal(STBorder.NIL);
							}

							if(stcb.getRight()!=null){      //判断是否有左
								stcb.getRight().setVal(STBorder.NIL);
							}else
							{
								CTBorder right = stcb.addNewRight();
								right.setVal(STBorder.NIL);
							}

						}else {                 //没有的话新建

							CTTcBorders tcb = tcpr.addNewTcBorders();
							CTBorder left = tcb.addNewLeft();
							left.setVal(STBorder.NIL);
							CTBorder right = tcb.addNewRight();
							right.setVal(STBorder.NIL);
						}
					}
				}
				else if (i == (trs.size() - 1))             //最后一行只删除上左右
				{
					for (CTTc tc : tcs)
					{

						tcpr = tc.getTcPr();

						CTTcBorders stcb = tcpr.getTcBorders();                //获取到原修饰体

						if (stcb!=null)         //判断是否有
						{
							if(stcb.getLeft()!=null){       //判断是否有左
								stcb.getLeft().setVal(STBorder.NIL);
							}else
							{
								CTBorder left = stcb.addNewLeft();
								left.setVal(STBorder.NIL);
							}

							if(stcb.getRight()!=null){      //判断是否有左
								stcb.getRight().setVal(STBorder.NIL);
							}else
							{
								CTBorder right = stcb.addNewRight();
								right.setVal(STBorder.NIL);
							}

							if(stcb.getTop()!=null){      //判断是否有上
								stcb.getTop().setVal(STBorder.NIL);
							}else
							{
								CTBorder top = stcb.addNewTop();
								top.setVal(STBorder.NIL);
							}

						}else
						{                 //没有的话新建

							CTTcBorders tcb = tcpr.addNewTcBorders();
							CTBorder top = tcb.addNewTop();
							top.setVal(STBorder.NIL);
							CTBorder left = tcb.addNewLeft();
							left.setVal(STBorder.NIL);
							CTBorder right = tcb.addNewRight();
							right.setVal(STBorder.NIL);
						}
					}
				}
				else for (CTTc tc : tcs)                    //其他全部处理
					{
						tcpr = tc.getTcPr();

						CTTcBorders stcb = tcpr.getTcBorders();                //获取到原修饰体

						if (stcb!=null)         //判断是否有
						{
							if(stcb.getLeft()!=null){       //判断是否有左
								stcb.getLeft().setVal(STBorder.NIL);
							}else
							{
								CTBorder left = stcb.addNewLeft();
								left.setVal(STBorder.NIL);
							}

							if(stcb.getRight()!=null){      //判断是否有左
								stcb.getRight().setVal(STBorder.NIL);
							}else
							{
								CTBorder right = stcb.addNewRight();
								right.setVal(STBorder.NIL);
							}

							if(stcb.getTop()!=null){      //判断是否有上
								stcb.getTop().setVal(STBorder.NIL);
							}else
							{
								CTBorder top = stcb.addNewTop();
								top.setVal(STBorder.NIL);
							}


							if(stcb.getBottom()!=null){      //判断是否有下
								stcb.getBottom().setVal(STBorder.NIL);
							}else
							{
								CTBorder bottom = stcb.addNewBottom();
								bottom.setVal(STBorder.NIL);
							}

						}else
						{                 //没有的话新建

							CTTcBorders tcb = tcpr.addNewTcBorders();
							CTBorder top = tcb.addNewTop();
							top.setVal(STBorder.NIL);
							CTBorder bottom = tcb.addNewBottom();
							bottom.setVal(STBorder.NIL);
							CTBorder left = tcb.addNewLeft();
							left.setVal(STBorder.NIL);
							CTBorder right = tcb.addNewRight();
							right.setVal(STBorder.NIL);
						}
					}

			}
		}
	}
}
