package Utils.WkHtmlToPdf;

import java.io.File;

public class WkHtmlToPdf {
    private static final String toPdfTool = "C:\\\\Program Files\\\\wkhtmltopdf\\\\bin\\\\wkhtmltopdf.exe";
    /**
     * 1.是先同过https://wkhtmltopdf.org/downloads.html下载WkHtmlToPdf这个软件
     * 2.用java调命令行完成转pdf的操作
     * 3.除了html转pdf外还能赚jpg
     * html转pdf
     * @param srcPath html路径，可以是硬盘上的路径，也可以是网络路径
     * @param destPath pdf保存路径
     * @return 转换成功返回true
     */
    public static boolean convert(String srcPath, String destPath){
        File file = new File(destPath);
        File parent = file.getParentFile();
        //如果pdf保存路径不存在，则创建路径
        if(!parent.exists()){
            parent.mkdirs();
        }
        StringBuilder cmd = new StringBuilder();
        cmd.append(toPdfTool);
        cmd.append(" ");
        cmd.append("  --header-line");//页眉下面的线
        cmd.append("  --header-center 这里是页眉这里是页眉这里是页眉这里是页眉 ");//页眉中间内容
        //cmd.append("  --margin-top 30mm ");//设置页面上边距 (default 10mm)
        cmd.append(" --header-spacing 10 ");//    (设置页眉和内容的距离,默认0)
        cmd.append(srcPath);
        cmd.append(" ");
        cmd.append(destPath);
        boolean result = true;
        try{
            Process proc = Runtime.getRuntime().exec(cmd.toString());
            HtmlToPdfInterceptor error = new HtmlToPdfInterceptor(proc.getErrorStream());
            HtmlToPdfInterceptor output = new HtmlToPdfInterceptor(proc.getInputStream());
            error.start();
            output.start();
            proc.waitFor();
        }catch(Exception e){
            result = false;
            e.printStackTrace();
        }
        return result;
    }

    public static void main(String[] args) {
        //file:///D:/archives/home/resource/imageRoot/by/0/2017/%E5%A4%A7%E4%BA%8B%E8%AE%B0/%E6%B5%8B%E8%AF%95%E4%B8%A4%E7%A7%8D%E6%96%B9%E6%B3%95%E7%94%9F%E6%88%90pdf2222.html
        String srcPath = "D:\\archives\\home\\resource\\imageRoot\\by\\0\\2017\\大事记\\测试两种方法生成pdf2222.html";
        String destPath = "D:/pdftext/wkhtmltopdf3.pdf";
        WkHtmlToPdf.convert(srcPath, destPath);
    }

}
