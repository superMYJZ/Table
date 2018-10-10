import java.math.BigInteger;
import java.util.ArrayList;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 *
 */

/**
 * @author Admin	My Paragraph
 *
 */
public class MP
{
    private int bT=0;								//几级标题 0代表是正文 不是标题
    private BigInteger szCs=new BigInteger("0");	//大小
    private String text=null;						//标题文本
    private ArrayList<XWPFParagraph> ps=new ArrayList<XWPFParagraph>();	//下面与第一个格式相同的p
    private String cs=null;							//字体
    private String eastAsia=null;					//是否黑体
    private int bE;                                 //IBodyElement所指向的数字就是在IBodyElements中的位置

    public MP() {}
    public MP(MP mp)
    {
        this.setbT(mp.getbT());
        this.setPs(mp.getPs());
        this.setSzCs(mp.getSzCs());
        this.setText(mp.getText());
        this.setCs(mp.getCs());
        this.setEastAsia(mp.getEastAsia());
        this.setbE(mp.getbE());
    }
    public MP(ArrayList<XWPFParagraph> p, String text,int bT,String cs,String eastAsia,int bE)
    {
        this.ps=p;
        this.text=text;
        this.bT=bT;
        this.cs=cs;
        this.eastAsia=eastAsia;
        this.bE=bE;
    }
    public void clear()	//初始化 恢复
    {
        this.bT=0;
        this.szCs=new BigInteger("0");
        this.text=null;
        this.ps=new ArrayList<XWPFParagraph>();
        this.cs=null;
        this.eastAsia=null;
        this.bE=0;
    }
    String getbTI() { return "toc "+bT;}
    int getbT()
    {
        return bT;
    }
    public void setbT(int bT)
    {
        this.bT=bT;
    }
    public BigInteger getSzCs()
    {
        return szCs;
    }
    public void setSzCs(BigInteger szCs)
    {
        this.szCs = szCs;
    }
    public String getText()
    {
        return text;
    }
    public void setText(String text)
    {
        this.text = text;
    }
    public ArrayList<XWPFParagraph> getPs()
    {
        return ps;
    }
    public void setPs(ArrayList<XWPFParagraph> ps)
    {
        this.ps = ps;
    }
    public void addPs(XWPFParagraph p)
    {
        this.ps.add(p);
    }
    public String getCs()
    {
        return cs;
    }
    public void setCs(String cs)
    {
        this.cs = cs;
    }
    public String getEastAsia()
    {
        return eastAsia;
    }
    public void setEastAsia(String eastAsia)
    {
        this.eastAsia = eastAsia;
    }
    public int getbE() {return bE;}
    public void setbE(int bE) {this.bE = bE;}

}
