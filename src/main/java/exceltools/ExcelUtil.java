package exceltools;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Author: georgexie
 * @Date: 2019/5/9 0009 14:20
 * @Version 1.0
 */
public class ExcelUtil {
        /**
         *
         * @Title: readXls
         * @Description: ����xls�ļ�
         * @param @param path
         * @param @return
         * @param @throws Exception    �趨�ļ�
         * @return List<List<String>>    ��������
         * @throws
         *
         * �Ӵ��벻�ѷ����䴦���߼���
         * 1.����InputStream��ȡexcel�ļ���io��
         * 2.Ȼ�󴩼�һ���ڴ��е�excel�ļ�HSSFWorkbook���Ͷ�����������ʾ������excel�ļ���
         * 3.�����excel�ļ���ÿҳ��ѭ������
         * 4.��ÿҳ��ÿ����ѭ������
         * 5.��ÿ���е�ÿ����Ԫ����������ȡ�����Ԫ���ֵ
         * 6.�����еĽ����ӵ�һ��List������
         * 7.��ÿ�еĽ����ӵ������ܽ����
         * 8.�������Ժ�ͻ�ȡ��һ��List<List<String>>���͵Ķ�����
         *
         */
        public List<List<String>> readXls(String path) throws Exception {
            InputStream is = new FileInputStream(path);
            // HSSFWorkbook ��ʶ����excel
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            List<List<String>> result = new ArrayList<List<String>>();
            int size = hssfWorkbook.getNumberOfSheets();
            // ѭ��ÿһҳ��������ǰѭ��ҳ
            for (int numSheet = 0; numSheet < size; numSheet++) {
                // HSSFSheet ��ʶĳһҳ
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                // ����ǰҳ��ѭ����ȡÿһ��
                for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                    // HSSFRow��ʾ��
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    int minColIx = hssfRow.getFirstCellNum();
                    int maxColIx = hssfRow.getLastCellNum();
                    List<String> rowList = new ArrayList<String>();
                    // �������У���ȡ����ÿ��cellԪ��
                    for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                        // HSSFCell ��ʾ��Ԫ��
                        HSSFCell cell = hssfRow.getCell(colIx);
                        if (cell == null) {
                            continue;
                        }
                        rowList.add(getStringVal(cell));
                    }
                    result.add(rowList);
                }
            }
            return result;
        }

        /**
         *
         * @Title: readXlsx
         * @Description: ����Xlsx�ļ�
         * @param @param path
         * @param @return
         * @param @throws Exception    �趨�ļ�
         * @return List<List<String>>    ��������
         * @throws
         */
        public List<List<String>> readXlsx(String path) throws Exception {
            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            List<List<String>> result = new ArrayList<List<String>>();
            // ѭ��ÿһҳ��������ǰѭ��ҳ
            for (XSSFSheet xssfSheet : xssfWorkbook) {
                if (xssfSheet == null) {
                    continue;
                }
                // ����ǰҳ��ѭ����ȡÿһ��
                for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    int minColIx = xssfRow.getFirstCellNum();
                    int maxColIx = xssfRow.getLastCellNum();
                    List<String> rowList = new ArrayList<String>();
                    for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                        XSSFCell cell = xssfRow.getCell(colIx);
                        if (cell == null) {
                            continue;
                        }
                        rowList.add(cell.toString());
                    }
                    result.add(rowList);
                }
            }
            return result;
        }

        // ���ڵ�����
    /*
     * ��ʵ��ʱ������ϣ���õ������ݾ���excel�е����ݣ���������ֽ��������
     * ������excel�е����������֣���ᷢ��Java�ж�Ӧ�ı���˿�ѧ��������
     * �����ڻ�ȡֵ��ʱ���Ҫ��һЩ���⴦������֤�õ��Լ���Ҫ�Ľ��
     * ���ϵ������Ƕ�����ֵ�͵����ݸ�ʽ������ȡ�Լ���Ҫ�Ľ����
     * �����ṩ����һ�ַ������ڴ�֮ǰ�������ȿ�һ��poi�ж���toString()����:
     *
     * �÷�����poi�ķ�������Դ�������ǿ��Է��֣��ô��������ǣ�
     * 1.��ȡ��Ԫ�������
     * 2.�������͸�ʽ�����ݲ�����������Ͳ����˺ܶ಻��������Ҫ��
     * �ʶ����������һ�����졣
     */
    /*public String toString(){
        switch(getCellType()){
            case CELL_TYPE_BLANK:
                return "";
            case CELL_TYPE_BOOLEAN:
                return getBooleanCellValue() ? "TRUE" : "FALSE";
            case CELL_TYPE_ERROR:
                return ErrorEval.getText(getErrorCellValue());
            case CELL_TYPE_FORMULA:
                return getCellFormula();
            case CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(this)){
                    DateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy")
                    return sdf.format(getDateCellValue());
                }
                return getNumericCellValue() + "";
            case CELL_TYPE_STRING:
                return getRichStringCellValue().toString();
            default :
                return "Unknown Cell Type:" + getCellType();
        }
    }*/

        /**
         * ����poiĬ�ϵ�toString������������
         * @Title: getStringVal
         * @Description: 1.���ڲ���Ϥ�����ͣ�����Ϊ���򷵻�""���ƴ�
         *               2.��������֣����޸ĵ�Ԫ������ΪString��Ȼ�󷵻�String�������ͱ�֤���ֲ�����ʽ����
         * @param @param cell
         * @param @return    �趨�ļ�
         * @return String    ��������
         * @throws
         */
        public static String getStringVal(HSSFCell cell) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                case Cell.CELL_TYPE_FORMULA:
                    return cell.getCellFormula();
                case Cell.CELL_TYPE_NUMERIC:
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    return cell.getStringCellValue();
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                default:
                    return "";
            }
        }

}
