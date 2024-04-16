package com.example.demoapachepoi.demoExportPDF.service;

import com.example.demoapachepoi.demoExportPDF.entity.MyObject;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Type;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class DecisionService {

    String jsonString = "[" +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thành Trung\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Tran Ngoc Anh\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Pham Tuan Tu\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thi\n Thuy\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Thiếu tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Lê Văn Giang\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Trung úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Tiến Tùng\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}]" +
            "]";

    Map<String, Object> mapData = new HashMap<>();

    @SneakyThrows
    public void export() {
        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);

            replaceText(doc, "${DangUyRQD}", "Đảng bộ tập đoàn Viettel");
            replaceText(doc, "${DangBoCapTren}", "Đảng bộ tập đoàn FPT");
            replaceText(doc, "${TenCBMoi}", "Chi bộ GPS");
            replaceText(doc, "${DangCapTrenTrucTiep}", "Đảng bộ tập đoàn NB");
            replaceText(doc, "${NhiemKy}", "2020-2022");
            replaceText(doc, "${CoQuanSoanThao}", "FPT");
            replaceText(doc, "${CacDieuTren}", "1,2,3"); //todo: thêm stt các điều vào đây
            replaceText(doc, "${SapNhap}", "trên cơ sở sáp nhập Chi bộ Team A và Chi bộ Team B");
            replaceText(doc, "${SoLuongDangVien}", "6");
            replaceText(doc, "${ChiUyMoi}", "Chi ủy GPS");
            replaceText(doc, "${DieuCuoiCung}", "4");
            replaceText(doc, "${ChuCaiDauTien}", "T");
            replaceText(doc, "${SoLuongBanPhatHanh}", "7");
            replaceText(doc, "${DiaDanh}", "Hà Nội");
            replaceText(doc, "${partyMemberList", jsonString);
//            returnText(doc);

            // Save the modified document
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output.docx");
            doc.write(out);

            // Close resources
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("Report generated successfully!");
            doc.close();
        }
    }

    public Map<String, Object> returnText(XWPFDocument document){

        XWPFRun run, nextRun;
        StringBuilder text = null;
        for (int a = 0; a < document.getParagraphs().size(); a++) {
            XWPFParagraph paragraph = document.getParagraphs().get(a);
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);
                    if(text != null && !checkSymmetric(text.toString())){
                        text.append("\n").append(run.getText(0));
                    } else {
                        text = new StringBuilder(run.getText(0));
                    }

                    if (text == null) {
                        continue;
                    }
                    if (checkVariable(text.toString()) || text.toString().contains("$")) {
                        while (i < runs.size() - 1) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            paragraph.removeRun(i + 1);
                            if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                break;
                            }
                        }
                        if(!checkSymmetric(text.toString())){
//                            run.setText("", 0);
//                            document.getParagraphs().remove(a);
                            continue;
                        }

                        mapData.put(getKey(text.toString()), getValue(text.toString()));
//                        run.setText(getValue(text.toString()), 0);
                    }
                }
            }
        }

        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                if(text != null && !checkSymmetric(text.toString())){
                                    text.append("\n").append(run.getText(0));
                                } else {
                                    text = new StringBuilder(run.getText(0));
                                }

                                if (text == null) {
                                    continue;
                                }
                                if (checkVariable(text.toString()) || text.toString().contains("$")) {
                                    while (i < runs.size() - 1) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        paragraph.removeRun(i + 1);
                                        if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                            break;
                                        }
                                    }
                                    if(!checkSymmetric(text.toString())){
//                                        run.setText("", 0);
                                        continue;
                                    }

                                    mapData.put(getKey(text.toString()), getValue(text.toString()));
//                                    run.setText(getValue(text.toString()), 0);
                                }
                            }
                        }
                    }
                }
            }
        }
        return mapData;
    }

    public void replaceTextFromDataCustom(XWPFDocument document){
        XWPFRun run, nextRun;
        StringBuilder text;
        for (XWPFParagraph paragraph: document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);
                    text = new StringBuilder(run.getText(0));
                    if (text == null) {
                        continue;
                    }
                    if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                        while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            paragraph.removeRun(i + 1);
                        }

                        String key = getKey(text.toString());
                        String value = mapData.get(key).toString();

                        //nếu thấy \n thì xuống dòng
                        if(value.contains("\n")){
                            String [] dataList = value.split("\n");
                            if(text.toString().contains(key)){
                                for (String data : dataList) {
                                    run.setText("", 0);
                                    XWPFRun newRun = paragraph.createRun();
                                    newRun.setText(data, 0);
                                    newRun.setFontSize(run.getFontSize());
                                    newRun.addCarriageReturn();
                                }
                            }
                        } else {
                            run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                        }
                    }
                }
            }
        }

        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
//                forTable(row.getTable(), key, value);
                for (XWPFTableCell cell : row.getTableCells()) {
//                    forTables(cell.getTables(), key, value);
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                text = new StringBuilder(run.getText(0));
                                if (text == null) {
                                    continue;
                                }
                                if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                                    while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        paragraph.removeRun(i + 1);
                                    }

//                                    if(text.toString().contains(key) && (key.contains("List") || key.contains("list"))){
//                                        value = replaceList(text.toString(), value);
//                                    }
//
//                                    //nếu thấy \n thì xuống dòng
//                                    if(value.contains("\n")){
//                                        String [] dataList = value.split("\n");
//                                        if(text.toString().contains(key)){
//                                            for (String data : dataList) {
//                                                run.setText("", 0);
//                                                XWPFRun newRun = paragraph.createRun();
//                                                newRun.setText(data, 0);
//                                                newRun.setFontSize(run.getFontSize());
//                                                newRun.addCarriageReturn();
//                                            }
//                                        }
//                                    } else {
//                                        run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
//                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private void replaceText(XWPFDocument document, String key, String value) {
        XWPFRun run, nextRun;
        StringBuilder text;
        for (XWPFParagraph paragraph: document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);
                    text = new StringBuilder(run.getText(0));
                    if (text == null) {
                        continue;
                    }
                    if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                        while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            paragraph.removeRun(i + 1);
                        }

                        if(text.toString().contains(key) && (key.contains("List") || key.contains("list"))){
                            value = replaceList(text.toString(), value);
                        }

                        //nếu thấy \n thì xuống dòng
                        if(value.contains("\n")){
                            String [] dataList = value.split("\n");
                            if(text.toString().contains(key)){
                                for (String data : dataList) {
                                    run.setText("", 0);
                                    XWPFRun newRun = paragraph.createRun();
                                    newRun.setText(data, 0);
                                    newRun.setFontSize(run.getFontSize());
//                                    newRun.addBreak();

//                                    paragraphNew.setIndentationLeft(0);
//                                    paragraphNew.setIndentationFirstLine(0);
                                    newRun.addCarriageReturn();
                                }
                            }
                        } else {
                            run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                        }
                    }
                }
            }
        }

        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
//                forTable(row.getTable(), key, value);
                for (XWPFTableCell cell : row.getTableCells()) {
//                    forTables(cell.getTables(), key, value);
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                text = new StringBuilder(run.getText(0));
                                if (text == null) {
                                    continue;
                                }
                                if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                                    while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        paragraph.removeRun(i + 1);
                                    }

                                    if(text.toString().contains(key) && (key.contains("List") || key.contains("list"))){
                                        value = replaceList(text.toString(), value);
                                    }

                                    //nếu thấy \n thì xuống dòng
                                    if(value.contains("\n")){
                                        String [] dataList = value.split("\n");
                                        if(text.toString().contains(key)){
                                            for (String data : dataList) {
                                                run.setText("", 0);
                                                XWPFRun newRun = paragraph.createRun();
                                                newRun.setText(data, 0);
                                                newRun.setFontSize(run.getFontSize());
//                                    newRun.addBreak();

//                                    paragraphNew.setIndentationLeft(0);
//                                    paragraphNew.setIndentationFirstLine(0);
                                                newRun.addCarriageReturn();
                                            }
                                        }
                                    } else {
                                        run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }


    private void forTable (XWPFTable tbl, String key, String value) {
        XWPFRun run, nextRun;
        String runsText;
        for (XWPFTableRow row : tbl.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    List<XWPFRun> runs = p.getRuns();
                    if (runs != null) {
                        for (int i =
                             0; i < runs.size(); i++) {
                            run = runs.get(i);
                            runsText = run.getText(0);
                            if (runsText == null) {
                                continue;
                            }
                            if (runsText.contains("${") || (runsText.contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                                while (!runsText.contains("}")) {
                                    nextRun = runs.get(i + 1);
                                    runsText = runsText + nextRun.getText(0);
                                    p.removeRun(i + 1);
                                }
                                run.setText(runsText.contains(key) ? runsText.replace(key, value) : runsText, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    private void forTables (List<XWPFTable> ltbl, String key, String value) {
        XWPFRun run, nextRun;
        String runsText;
        for(XWPFTable tbl : ltbl) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        List<XWPFRun> runs = p.getRuns();
                        if (runs != null) {
                            for (int i =
                                 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                runsText = run.getText(0);
                                if (runsText == null) {
                                    continue;
                                }
                                if (runsText.contains("${") || (runsText.contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                                    while (!runsText.contains("}")) {
                                        nextRun = runs.get(i + 1);
                                        runsText = runsText + nextRun.getText(0);
                                        p.removeRun(i + 1);
                                    }
                                    run.setText(runsText.contains(key) ? runsText.replace(key, value) : runsText, 0);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private String replaceList(String textTemplate, String partyMembers){
        textTemplate = removeKeyParamList(textTemplate);
        //get value by key: lấy ra được nội dung giống như 'partyMemberList'
        Gson gson = new Gson();
        Type listType = new TypeToken<List<List<MyObject>>>(){}.getType();
        List<List<MyObject>> myObjects = gson.fromJson(partyMembers, listType);
        StringBuilder finalText = new StringBuilder();
        for (int j = 0; j < myObjects.size(); j++){
            String text = "";
            String [] lst = textTemplate.split(" ");
            for (int i = 0; i < lst.length; i++){
                for(MyObject myObject : myObjects.get(j)){
                    if(lst[i].contains("${")){
                        if(lst[i].contains("STT")){
                            lst[i] = String.valueOf(j + 1).concat(".");
                            break;
                        }
                        if(lst[i].contains(myObject.getKey())){
                            lst[i] = myObject.getValue();
                            if (lst[i].contains(".")){
                                lst[i] += ".";
                            }
                            break;
                        }
                    }
                }
                text = String.join(" ", lst);
            }
            finalText.append(text);
            finalText.append("\n");
        }
        return finalText.toString();
    }

    public static boolean checkSymmetric(String s) {
        Stack<Character> stack = new Stack<>();
        for (char ch : s.toCharArray()) {
            if (ch == '{') {
                stack.push(ch);
            } else if (ch == '}') {
                if (stack.isEmpty() || stack.pop() != '{') {
                    return false;
                }
            }
        }
        return stack.isEmpty();
    }

    public String removeKeyParamList(String key){
        int start = key.indexOf("${", 2);
        int finish = key.lastIndexOf("}");
        return key.substring(start, finish);
    }

    public boolean checkVariable(String text){
        String regex = "\\$[^{]*\\{";
        Pattern pattern = Pattern.compile(regex);
        boolean b = pattern.matcher(text).find();
        return b;
    }

    public String getKey(String text){
        String regex = "\\$([^\\{]*)\\{";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            String key = matcher.group(1);
            return key;
        }
        return null;
    }

    public String getValue(String text){
        String regex = "\\{([^}]*)\\}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            String value = matcher.group(1);
            return value;
        }
        return null;
    }

}
