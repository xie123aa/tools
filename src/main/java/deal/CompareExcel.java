package deal;

import exceltools.ExcelUtil;

import java.util.List;

/**
 * @Author: georgexie
 * @Date: 2019/5/9 0009 14:40
 * @Version 1.0
 */
public class CompareExcel {
    public static void main(String[] args) {
        String path = "C:\\Users\\Administrator\\Desktop\\补数据\\20190508导入数据.xls";
        try {
            List<List<String>> result = new ExcelUtil().readXls(path);
            System.out.println(result.size());
            for (int i = 0; i < result.size(); i++) {
                List<String> model = result.get(i);
                System.out.println("orderNum:" + model.get(0) + "--> orderAmount:" + model.get(1));
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
