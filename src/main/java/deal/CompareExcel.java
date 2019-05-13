package deal;

import exceltools.ExcelUtil;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Author: georgexie
 * @Date: 2019/5/9 0009 14:40
 * @Version 1.0
 */
public class CompareExcel {
    public static void main(String[] args) {
    String filePath = "C:\\������\\�����\\20190508��������.xlsx";
    Workbook wb =null;
    Sheet sheet = null;
    Row row = null;
    List<Map<String,String>> list = null;
    String cellData = null;
    String columns[] = { "������Ϣ��ˮ��","������ˮ��","������ˮ��","������ˮ��","ת����","������ˮ��",
            "�û�����","��ϵ�绰","�����漰����","ͨѶ��ַ","������λ","��������","Ͷ�߱���","��Ͷ��ҵ",
            "�漰��ҵ","����ʡ��","�ؼ���","������(һ)","������(��)","������(��)","ҵ����(һ)","ҵ����(��)",
            "ҵ����(��)","������Դ","��������","�϶�Ա","����Ա","����Ա","����ʽ","����״̬","�϶����","����ͽ���","ת��ͽ���","�����ͽ���","�Ƿ񲢰�",
            "��������","��������","ת������","��������","�᰸����","Ԥ��������","��������",
            "��֤Ա���","������ԭ��","�Զ�����ע","������","�������","","","",""};
    wb = readExcel(filePath);
        if(wb != null){
        //������ű�������
        list = new ArrayList<Map<String,String>>();
        //��ȡ��һ��sheet
        sheet = wb.getSheetAt(0);
        //��ȡ�������
        int rownum = sheet.getPhysicalNumberOfRows();
        //��ȡ��һ��
        row = sheet.getRow(0);
        //��ȡ�������
        int colnum = row.getPhysicalNumberOfCells();
        for (int i = 1; i<rownum; i++) {
            Map<String,String> map = new LinkedHashMap<String,String>();
            row = sheet.getRow(i);
            if(row !=null){
                for (int j=0;j<colnum;j++){
                    cellData = (String) getCellFormatValue(row.getCell(j));
                    map.put(columns[j], cellData);
                }
            }else{
                break;
            }
            list.add(map);
        }
    }
    //��������������list
        for (Map<String,String> map : list) {
        for (Map.Entry<String,String> entry : map.entrySet()) {
            System.out.print(entry.getKey()+":"+entry.getValue()+",");
        }
        System.out.println();
    }

}
    //��ȡexcel
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //�ж�cell����
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //�ж�cell�Ƿ�Ϊ���ڸ�ʽ
                    if(DateUtil.isCellDateFormatted(cell)){
                        //ת��Ϊ���ڸ�ʽYYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //����
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }

}
