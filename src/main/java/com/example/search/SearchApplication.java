package com.example.search;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.net.ssl.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.KeyManagementException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.X509Certificate;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Getter
@Setter
@Data
class Vo {
    private String productName;

    private Integer price;

    private String style;

    private String link;
}


@SpringBootApplication
public class SearchApplication {

    public static void main(String[] args) throws IOException {
        //ssh
        disableCertificateValidation();


        List<Vo> dataList = new ArrayList<>();

        Scanner sc = new Scanner(System.in);
        System.out.println("請輸入要查詢的關鍵字");
        String keyWord = sc.next();

        // maxPage
        Document doc = Jsoup.connect("https://m.momoshop.com.tw/search.momo?searchKeyword=" + keyWord).get();

        // 取得 <head> 標籤的内容
        String headContent = doc.selectFirst("head").html();

        // 取得頁面最大值的值
        Matcher maxPageMatcher = Pattern.compile("var maxPage\\s*=\\s*'([\\d]+)'").matcher(headContent);
        String maxPage = null;
        if (maxPageMatcher.find()) {
            maxPage = maxPageMatcher.group(1);
            System.out.println("maxPage value: " + maxPage);
        } else {
            System.out.println("maxPage value not found.");
        }

        for (int i = 1; i < Integer.parseInt(maxPage); i++) {

            // 建構網址
            StringBuilder netUrl = new StringBuilder();
            netUrl.append("https://m.momoshop.com.tw/search.momo").append("?searchKeyword=").append(keyWord).append("&cpCode=&couponSeq=&searchType=1&cateLevel=-1&curPage=").append(i).append("&cateCode=&cateName=&maxPage=").append(maxPage).append("&minPage=1&_advCp=N&_advFirst=N&_advFreeze=N&_advSuperstore=N&_advTvShop=N&_advTomorrow=N&_advNAM=N&_advStock=N&_advPrefere=N&_advThreeHours=N&_advVideo=N&_advCycle=N&_advCod=N&_advSuperstorePay=N&_advPriceS=&_advPriceE=&_brandNameList=&_brandNoList=&brandSeriesStr=&isBrandSeriesPage=0&ent=b&_imgSH=fourCardType&specialGoodsType=&_isFuzzy=0&_spAttL=&_mAttL=&_sAttL=&_noAttL=&topMAttL=&topSAttL=&topNoAttL=&hotKeyType=0&hashTagCode=&hashTagName=");

            // 送HTTP請求並設置header
            Connection connection = Jsoup.connect(netUrl.toString()).timeout(120000).maxBodySize(0);
            connection.userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36");

            // 資料處理
            try {
                Document pageDoc = connection.get();
                Elements prdName = pageDoc.select("body ul li a div div h3.prdName");
                Elements prices = pageDoc.select("body ul li a div p span b[class=price ]");
                Elements links = pageDoc.select("body article.prdListArea ul li a");

                for (int n = 0; n < prdName.size(); n++) {
                    Vo data = new Vo();
                    Element name = prdName.get(n);
                    data.setProductName(getString(name.toString()));

                    // 檢查字串是否符合以連字符號-前後為字母或數字的正則表達式
                    String regex = "\\b([A-Za-z0-9]+-[A-Za-z0-9]+)\\b";
                    Pattern pattern = Pattern.compile(regex);
                    Matcher matcher = pattern.matcher(name.toString());

                    if (matcher.find()) {
                        String extractedString = matcher.group(1);
                        data.setStyle(extractedString);
                    }

                    // 處理價格
                    if (n < prices.size()) {
                        Element price = prices.get(n);
                        data.setPrice(Integer.parseInt(getString(price.toString())));
                    }

                    // 處理連結
                    if (n < links.size()) {
                        String link = links.get(n).attr("href");
                        // 檢查字串是否符合以 /goods.momo 為開頭的正則表達式
                        String regexLink = "^/goods\\.momo.*";
                        boolean match = Pattern.matches(regexLink, link);
                        if (match) {
                            data.setLink("https://m.momoshop.com.tw/" + link);
                        }
                    }
                    dataList.add(data);
                }




            } catch (IOException e) {
                e.printStackTrace();
            }
        }


        System.out.println("dataList=" + dataList);


        // 建立 Excel 工作簿
        Workbook workbook = new XSSFWorkbook();
        // 建立工作表
        Sheet sheet = workbook.createSheet("Product List");

        // 建立標題行
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("品名");
        headerRow.createCell(1).setCellValue("價格");
        headerRow.createCell(2).setCellValue("型號");
        headerRow.createCell(3).setCellValue("網址");

        // 寫入資料
//            int rowNum = 1;
//            for (Map.Entry<String, String> entry : product.entrySet()) {
//                Row row = sheet.createRow(rowNum++);
//                row.createCell(0).setCellValue(entry.getKey());
//                row.createCell(1).setCellValue(entry.getValue());
//            }
        int rowNum = 1;
        for (Vo vo : dataList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(vo.getProductName());
            row.createCell(1).setCellValue(vo.getPrice());
            row.createCell(2).setCellValue(vo.getStyle());
            row.createCell(3).setCellValue(vo.getLink());
        }

        // 取得當前日期並格式化為特定的字串形式
        LocalDate currentDate = LocalDate.now();
        String dateString = currentDate.format(DateTimeFormatter.ofPattern("yyyyMMdd"));

        // 設定下載檔案的相關設定
        String filename = dateString + ".xlsx";
        String downloadPath = "C:\\downloadfile\\" + filename;

        // 寫入 Excel 檔案
        try (FileOutputStream outputStream = new FileOutputStream(downloadPath)) {
            workbook.write(outputStream);
        }

        System.out.println("Excel 檔案寫入完成");

        // 下載 Excel 檔案
        System.out.println("下載 Excel 檔案至本地端：" + downloadPath);
        // 在此處實現將 Excel 檔案下載到本地端的程式碼


        System.out.println("執行完畢，請關閉視窗");
    }

    private static String getString(String html) {
        // 選取文字部分的起始位置和結束位置
        int startIndex = html.indexOf(">") + 1;
        int endIndex = html.lastIndexOf("<");
        // 提取文字
        String text = html.substring(startIndex, endIndex);

        return text;
    }

    // 允許連接到任何HTTPS服務器而不進行證書和主機名的驗證
    private static void disableCertificateValidation() {
        TrustManager[] trustAllCerts = new TrustManager[]{new X509TrustManager() {
            //        返回null表示該信任管理器不提供任何可接受的發行者
            public java.security.cert.X509Certificate[] getAcceptedIssuers() {
                return null;
            }

            public void checkClientTrusted(X509Certificate[] certs, String authType) {
            }

            public void checkServerTrusted(X509Certificate[] certs, String authType) {
            }
        }};

        try {
            SSLContext sslContext = SSLContext.getInstance("SSL");
            sslContext.init(null, trustAllCerts, new java.security.SecureRandom());
            HttpsURLConnection.setDefaultSSLSocketFactory(sslContext.getSocketFactory());

            HostnameVerifier allHostsValid = (hostname, session) -> true;

            HttpsURLConnection.setDefaultHostnameVerifier(allHostsValid);
        } catch (NoSuchAlgorithmException | KeyManagementException e) {
            e.printStackTrace();
        }
    }


}

