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
import java.util.List;
import java.util.Stack;

@Service
public class DecisionService {

    private String partyMemberList = "${partyMemberList ${STT}. Đồng chí ${QuanHam} ${HoTen} - ${ChucVuChinhQuyen} - ${ChucVuCapUy}}";

    String jsonString = "[" +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thành Trung\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Tran Ngoc Anh\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Pham Tuan Tu\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thi\n Thuy\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Thiếu tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Lê Văn Giang\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Trung úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Tiến Tùng\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}]" +
            "]";

    @SneakyThrows
    public void export() {
        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);

            replaceText(doc, "${DangUyRQD}", "Đảng bộ tập đoàn Viettel");
            replaceText(doc, "${DangBoCapTren}", "Đảng bộ tập đoàn FPT");
            replaceText(doc, "${partyMemberList", jsonString);

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
                        while (!text.toString().contains("}") && !checkSymmetric(text.toString())) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            paragraph.removeRun(i + 1);
                        }

                        if(text.toString().contains(key)){
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
                forTable(row.getTable(), key, value);
                for (XWPFTableCell cell : row.getTableCells()) {
                    forTables(cell.getTables(), key, value);
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        List<XWPFRun> runs = p.getRuns();
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                text = new StringBuilder(run.getText(0));
                                if (text == null) {
                                    continue;
                                }
                                if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                                    while (!text.toString().contains("}")) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        p.removeRun(i + 1);
                                    }
                                    run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
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
                            lst[i] = String.valueOf(j + 1).concat(". ");
                            break;
                        }
                        if(lst[i].contains(myObject.getKey())){
                            lst[i] = myObject.getValue();
                            if (lst[i].contains(".")){
                                lst[i] += ". ";
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

}
