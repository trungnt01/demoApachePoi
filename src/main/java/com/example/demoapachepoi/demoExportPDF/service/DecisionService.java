package com.example.demoapachepoi.demoExportPDF.service;

import com.example.demoapachepoi.demoExportPDF.entity.MyObject;
import com.example.demoapachepoi.demoExportPDF.entity.PartyMember;
import com.example.demoapachepoi.demoExportPDF.utils.MultiValueMap;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class DecisionService {

    String jsonString = "[" +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thành Trung Nguyễn Thành Trung Nguyễn Thành Trung Nguyễn Thành Trung\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Tran Ngoc Anh Tran Ngoc Anh Tran Ngoc Anh Tran Ngoc Anh Tran Ngoc AnTran Ngoc Anh\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Pham Tuan Tu\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Thi\n Thuy\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Thiếu tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Lê Văn Giang\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Trung úy\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}], " +
                "[{\"key\":\"HoTen\",\"value\":\"Nguyễn Tiến Tùng\",\"access\":\"read_only\"},{\"key\":\"QuanHam\",\"value\":\"Đại tá\",\"access\":\"read_only\"},{\"key\":\"ChucVuChinhQuyen\",\"value\":\"Chủ tịch\",\"access\":\"read_only\"},{\"key\":\"ChucVuCapUy\",\"value\":\"Bí thư\",\"access\":\"read_only\"},{\"key\":\"tuoi\",\"value\":\"32\",\"access\":\"read_only\"},{\"key\":\"diaChi\",\"value\":\"Hà Nội\",\"access\":\"read_only\"}]" +
            "]";

    List<PartyMember> partyMembers = Arrays.asList(
            new PartyMember("Nguyễn Thành Trung Nguyễn Thành Trung Nguyễn Thành Trung", "Đại úy", "Trưởng phòng", "Bí thư")
            ,new PartyMember("Trần Văn Toàn", "Trung úy", "Phó phòng", "Phó bí thư")
            ,new PartyMember("Phạm Tiến Hùng", "Đại tá", "Tổ trưởng", "Ủy viên")
            ,new PartyMember("Lê Thị Hương Lê Thị Hương Lê Thị Hương Lê Thị Hương", "Thiếu úy", "Nhân viên", "Ủy viên")
            ,new PartyMember("Nguyễn Hương Giang", "Đại úy", "Nhân viên", "Ủy viên")
            ,new PartyMember("Trần Tuấn Hà", "Thượng tá", "Culi", "Ủy viên")
    );


    @SneakyThrows
    public Map<String, String> exportText() {
        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);

            replaceText(doc, "${DangBoCapTren}", "Đảng bộ tập đoàn FPT");
            replaceText(doc, "${DangUyRQD}", "Đảng bộ tập đoàn Viettel");
            replaceText(doc, "${partyMemberList", "jsonString");
            replaceText(doc, "${TenCBMoi}", "Chi bộ GPS");
            replaceText(doc, "${DangCapTrenTrucTiep}", "Đảng bộ tập đoàn NB");
            replaceText(doc, "${NhiemKy}", "2020-2022");
            replaceText(doc, "${CoQuanSoanThao}", "FPT");
            replaceText(doc, "${CacDieuTren}", "1,2,3"); //todo: thêm stt các điều vào đây
            replaceText(doc, "${SapNhap}", "trên cơ sở sáp nhập Chi bộ Team A và Chi bộ Team B");
            replaceText(doc, "${SoLuongDangVien}", "6");
            replaceText(doc, "${ChiUyMoi}", "Chi ủy GPS");
            replaceText(doc, "${DieuCuoiCung}", "4");
            replaceText(doc, "${ChuCaiDau}", "T");
            replaceText(doc, "${SoLuongBanPhatHanh}", "7");
            replaceText(doc, "${DiaDanh}", "Hà Nội");
            replaceText(doc, "${ChucDanhNguoiKy}", "Bí thư");


            Map<String, String> textMap = returnText(doc);
            Map<String, Map<String, String>> styleMap = returnStyle(doc);

//            replaceTextFromDataCustom();

            // Save the modified document
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output.docx");
            doc.write(out);
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("Report generated successfully!");
            doc.close();
            return textMap;
        }
    }
    @SneakyThrows
    public Map<String, Map<String, String>> exportStyle() {
        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);

            replaceText(doc, "${DangBoCapTren}", "Đảng bộ tập đoàn FPT");
            replaceText(doc, "${DangUyRQD}", "Đảng bộ tập đoàn Viettel");
            replaceText(doc, "${partyMemberList", "jsonString");
            replaceText(doc, "${TenCBMoi}", "Chi bộ GPS");
            replaceText(doc, "${DangCapTrenTrucTiep}", "Đảng bộ tập đoàn NB");
            replaceText(doc, "${NhiemKy}", "2020-2022");
            replaceText(doc, "${CoQuanSoanThao}", "FPT");
            replaceText(doc, "${CacDieuTren}", "1,2,3"); //todo: thêm stt các điều vào đây
            replaceText(doc, "${SapNhap}", "trên cơ sở sáp nhập Chi bộ Team A và Chi bộ Team B");
            replaceText(doc, "${SoLuongDangVien}", "6");
            replaceText(doc, "${ChiUyMoi}", "Chi ủy GPS");
            replaceText(doc, "${DieuCuoiCung}", "4");
            replaceText(doc, "${ChuCaiDau}", "T");
            replaceText(doc, "${SoLuongBanPhatHanh}", "7");
            replaceText(doc, "${DiaDanh}", "Hà Nội");
            replaceText(doc, "${ChucDanhNguoiKy}", "Bí thư");


            Map<String, String> textMap = returnText(doc);
            Map<String, Map<String, String>> styleMap = returnStyle(doc);

//            replaceTextFromDataCustom();

            // Save the modified document
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output.docx");
            doc.write(out);
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("Report generated successfully!");
            doc.close();
            return styleMap;
        }
    }


    //trả về style của từng run
    public Map<String, Map<String, String>> returnStyle(XWPFDocument document){
        Map<String, Map<String, String>> mapStyle2 = new LinkedHashMap<>();

        XWPFRun run, nextRun;
        StringBuilder text = null;

        for(IBodyElement iBodyElement : document.getBodyElements()){
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;

                Map<String, String> map = new LinkedHashMap<>();
                String textParagraph = paragraph.getText();

                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null && !runs.isEmpty()) {
                    for (int i = 0; i < runs.size(); i++) {
                        run = runs.get(i);
                        text = new StringBuilder(run.getText(0));
                        while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
                                && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":"))
                                && runs.size() > 1 && i < runs.size() - 1
                        ) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            i++;
                        }
                        boolean bold = run.isBold();
                        boolean italic = run.isItalic();
                        String color = run.getColor();
                        int fontSize = run.getFontSize();

                        StringBuilder styleBuilder = new StringBuilder();
                        styleBuilder.append("bold:").append(bold).append(",")
                                .append("italic:").append(italic).append(",")
                                .append("color:").append(color).append(",")
                                .append("fontSize:").append(fontSize)
                        ;
                        map.put(text.toString(), styleBuilder.toString());
                    }
                    mapStyle2.put(textParagraph, map);
                }

            } else if (iBodyElement instanceof XWPFTable) {
                XWPFTable tbl = (XWPFTable) iBodyElement;
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {

                            String textParagraph = paragraph.getText();
                            Map<String, String> map = new LinkedHashMap<>();

                            List<XWPFRun> runs = paragraph.getRuns();
                            if (runs != null && !runs.isEmpty()) {
                                for (int i = 0; i < runs.size(); i++) {
                                    run = runs.get(i);
                                    text = new StringBuilder(run.getText(0));
                                    while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
                                            && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":"))
                                            && runs.size() > 1 && i < runs.size() - 1
                                    ) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        i++;
                                    }
                                    boolean bold = run.isBold();
                                    boolean italic = run.isItalic();
                                    String color = run.getColor();
                                    int fontSize = run.getFontSize();

                                    StringBuilder styleBuilder = new StringBuilder();
                                    styleBuilder.append("bold:").append(bold).append(",")
                                            .append("italic:").append(italic).append(",")
                                            .append("color:").append(color).append(",")
                                            .append("fontSize:").append(fontSize)
                                    ;
                                    map.put(text.toString(), styleBuilder.toString());
                                }
                                mapStyle2.put(textParagraph, map);
                            }
                        }
                    }
                }
            }
        }

//        for (int a = 0; a < document.getParagraphs().size(); a++) {
//            XWPFParagraph paragraph = document.getParagraphs().get(a);
//            List<XWPFRun> runs = paragraph.getRuns();
//            if (runs != null) {
//                for (int i = 0; i < runs.size(); i++) {
//                    run = runs.get(i);
//                    text = new StringBuilder(run.getText(0));
////                    if(checkVariable(text.toString())){
////                        keyStyle = text.toString();
////                        i++;
////                        continue;
////                    }
////                    if (text.toString().contains("$")) {
////                        while (i < runs.size() - 1 && !checkVariable(text.toString())) {
////                            nextRun = runs.get(i + 1);
////                            text.append(nextRun.getText(0));
////                            i++;
////                        }
////                        keyStyle = text.toString();
////                    }
////                    else {
//                        while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
//                            && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":"))
//                            && runs.size() > 1 && i < runs.size() - 1
//                        ) {
//                            nextRun = runs.get(i + 1);
//                            text.append(nextRun.getText(0));
//                            i++;
////                            if(!nextRun.getText(0).equals("}")){
////                            }
////                            break;
//                        }
//                        boolean bold = run.isBold();
//                        boolean italic = run.isItalic();
//                        StringBuilder styleBuilder = new StringBuilder();
//                        styleBuilder.append("bold:").append(bold).append(",")
//                                .append("italic:").append(italic);
//                        mapStyle.put(text.toString(), styleBuilder.toString());
////                    }
//                }
//            }
//        }
//
//        for (XWPFTable tbl : document.getTables()) {
//            for (XWPFTableRow row : tbl.getRows()) {
//                for (XWPFTableCell cell : row.getTableCells()) {
//                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
//                        List<XWPFRun> runs = paragraph.getRuns();
//                        if (runs != null) {
//                            for (int i = 0; i < runs.size(); i++) {
//                                run = runs.get(i);
//                                text = new StringBuilder(run.getText(0));
//                                while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
//                                        && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":"))
//                                        && runs.size() > 1 && i < runs.size() - 1
//                                ) {
//                                    nextRun = runs.get(i + 1);
//                                    text.append(nextRun.getText(0));
//                                    i++;
//                                }
//                                boolean bold = run.isBold();
//                                boolean italic = run.isItalic();
//                                StringBuilder styleBuilder = new StringBuilder();
//                                styleBuilder.append("bold:").append(bold).append(",")
//                                        .append("italic:").append(italic);
//                                mapStyle.put(text.toString(), styleBuilder.toString());
//                            }
//                        }
//                    }
//                }
//            }
//        }
        return mapStyle2;
    }

    //xóa key xác định từng đoạn trong template
    public void deleteKey(XWPFDocument document){
        XWPFRun run, nextRun;
        StringBuilder text = null;
        for (int a = 0; a < document.getParagraphs().size(); a++) {
            XWPFParagraph paragraph = document.getParagraphs().get(a);
            List<XWPFRun> runs = paragraph.getRuns();
            Integer runDelete = null;
            Integer run2Delete = null;
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);
                    if(text != null && !checkSymmetric(text.toString())){
                        text.append(run.getText(0));
                    } else {
                        text = new StringBuilder(run.getText(0));
                    }
                    if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                        runDelete = i;
                    }
                    if (text.toString().contains("$") && run.getText(0).contains("$")) {
                        runDelete = 0;
                        while (i < runs.size() - 1) {
                            nextRun = runs.get(i + 1);
                            String oldText = text.toString();
                            text.append(nextRun.getText(0));
                            if(!checkVariable(text.toString())){
                                paragraph.removeRun(i + 1);
                            } else if(checkVariable(text.toString()) && !checkVariable(oldText)){
                                paragraph.removeRun(i + 1);
                            }else {
                                i++;
                                if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                    run2Delete = i;
                                    break;
                                }
                            }
                        }
                    }
                }
                if(run2Delete != null){
                    paragraph.removeRun(run2Delete);
                    run2Delete = null;
                }
                if(runDelete != null){
                    paragraph.removeRun(runDelete);
                    runDelete = null;
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
                        Integer runDelete = null;
                        Integer run2Delete = null;
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                if(text != null && !checkSymmetric(text.toString())){
                                    text.append(run.getText(0));
                                } else {
                                    text = new StringBuilder(run.getText(0));
                                }
                                if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                                    runDelete = i;
                                }
                                if (text.toString().contains("$") && run.getText(0).contains("$")) {
                                    runDelete = 0;
                                    while (i < runs.size() - 1) {
                                        nextRun = runs.get(i + 1);
                                        String oldText = text.toString();
                                        text.append(nextRun.getText(0));
                                        if(!checkVariable(text.toString())){
                                            paragraph.removeRun(i + 1);
                                        } else if(checkVariable(text.toString()) && !checkVariable(oldText)){
                                            paragraph.removeRun(i + 1);
                                        }else {
                                            i++;
                                            if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                                run2Delete = i;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            if(run2Delete != null){
                                paragraph.removeRun(run2Delete);
                                run2Delete = null;
                            }
                            if(runDelete != null){
                                paragraph.removeRun(runDelete);
                                runDelete = null;
                            }
                        }
                    }
                }
            }
        }
    }


    //trả về data sau khi gen quyết định
    public Map<String, String> returnText(XWPFDocument document){

        Map<String, String> mapData = new LinkedHashMap<>();

        XWPFRun run, nextRun;
        StringBuilder text = null;

        for(IBodyElement iBodyElement : document.getBodyElements()){
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (int i = 0; i < runs.size(); i++) {
                        run = runs.get(i);
                        if(text != null && !checkSymmetric(text.toString())){
                            text.append("\n").append(run.getText(0));
                        } else {
                            text = new StringBuilder(run.getText(0));
                        }
                        if (checkVariable(text.toString()) || text.toString().contains("$")) {
                            while (i < runs.size() - 1) {
                                nextRun = runs.get(i + 1);
                                text.append(nextRun.getText(0));
                                i++;
                                if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                    break;
                                }
                            }
                            if(!checkSymmetric(text.toString())){
                                continue;
                            }
                            mapData.put(getKey(text.toString()), getValue(text.toString()));
                        }
                    }
                }
            } else if (iBodyElement instanceof XWPFTable) {
                XWPFTable tbl = (XWPFTable) iBodyElement;
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
                                    if (checkVariable(text.toString()) || text.toString().contains("$")) {
                                        while (i < runs.size() - 1) {
                                            nextRun = runs.get(i + 1);
                                            text.append(nextRun.getText(0));
                                            i++;
                                            if(text.toString().contains("{") && checkSymmetric(text.toString())){
                                                break;
                                            }
                                        }
                                        if(!checkSymmetric(text.toString())){
                                            continue;
                                        }
                                        mapData.put(getKey(text.toString()), getValue(text.toString()));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


//        for (int a = 0; a < document.getParagraphs().size(); a++) {
//            XWPFParagraph paragraph = document.getParagraphs().get(a);
//            List<XWPFRun> runs = paragraph.getRuns();
//            if (runs != null) {
//                for (int i = 0; i < runs.size(); i++) {
//                    run = runs.get(i);
//                    if(text != null && !checkSymmetric(text.toString())){
//                        text.append("\n").append(run.getText(0));
//                    } else {
//                        text = new StringBuilder(run.getText(0));
//                    }
//                    if (checkVariable(text.toString()) || text.toString().contains("$")) {
//                        while (i < runs.size() - 1) {
//                            nextRun = runs.get(i + 1);
//                            text.append(nextRun.getText(0));
////                            paragraph.removeRun(i + 1);
//                            i++;
//                            if(text.toString().contains("{") && checkSymmetric(text.toString())){
//                                break;
//                            }
//                        }
//                        if(!checkSymmetric(text.toString())){
//                            continue;
//                        }
//                        mapData.put(getKey(text.toString()), getValue(text.toString()));
//                    }
//                }
//            }
//        }
//
//        for (XWPFTable tbl : document.getTables()) {
//            for (XWPFTableRow row : tbl.getRows()) {
//                for (XWPFTableCell cell : row.getTableCells()) {
//                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
//                        List<XWPFRun> runs = paragraph.getRuns();
//                        if (runs != null) {
//                            for (int i = 0; i < runs.size(); i++) {
//                                run = runs.get(i);
//                                if(text != null && !checkSymmetric(text.toString())){
//                                    text.append("\n").append(run.getText(0));
//                                } else {
//                                    text = new StringBuilder(run.getText(0));
//                                }
//                                if (checkVariable(text.toString()) || text.toString().contains("$")) {
//                                    while (i < runs.size() - 1) {
//                                        nextRun = runs.get(i + 1);
//                                        text.append(nextRun.getText(0));
////                                      paragraph.removeRun(i + 1);
//                                        i++;
//                                        if(text.toString().contains("{") && checkSymmetric(text.toString())){
//                                            break;
//                                        }
//                                    }
//                                    if(!checkSymmetric(text.toString())){
//                                        continue;
//                                    }
//                                    mapData.put(getKey(text.toString()), getValue(text.toString()));
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//        }
        deleteKey(document);
        return mapData;
    }


    //replace param bằng data sau khi sửa
//    @SneakyThrows
//    public void replaceTextFromDataCustom(){
//        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
//            XWPFDocument document = new XWPFDocument(inputStream);
//
//
////            -----------------------------------------------------------------
//            XWPFRun run, nextRun;
//            StringBuilder text = null;
//            for (XWPFParagraph paragraph: document.getParagraphs()) {
//                List<XWPFRun> runs = paragraph.getRuns();
//                if (runs != null) {
//                    for (int i = 0; i < runs.size(); i++) {
//                        run = runs.get(i);
//                        if(text != null && !checkSymmetric(text.toString())){
//                            text.append("\n").append(run.getText(0));
//                        } else {
//                            text = new StringBuilder(run.getText(0));
//                        }
//                        if (text == null) {
//                            continue;
//                        }
//                        if (checkVariable(text.toString()) || text.toString().contains("$")) {
//                            while (i < runs.size() - 1) {
//                                nextRun = runs.get(i + 1);
//                                text.append(nextRun.getText(0));
//                                paragraph.removeRun(i + 1);
//                                if(text.toString().contains("{") && checkSymmetric(text.toString())){
//                                    break;
//                                }
//                            }
//                            if(!checkSymmetric(text.toString())){
////                                while (!paragraph.getRuns().isEmpty()) {
////                                    List<IBodyElement> bodyElements = document.getBodyElements();
////                                    boolean delete = false;
////                                    for (int y = 0; y < bodyElements.size(); y++){
////                                        IBodyElement bodyElement = bodyElements.get(y);
////                                        if(bodyElement instanceof XWPFParagraph){
////                                            XWPFParagraph xwpfParagraph = (XWPFParagraph) bodyElement;
////                                            if(compareParagraph(xwpfParagraph, paragraph)){
////                                                document.removeBodyElement(y);
////                                                delete = true;
////                                                break;
////                                            }
////                                        }
////                                    }
////                                    if(delete) break;
////                                }
//                                paragraph.removeRun(0);
//                                continue;
//                            }
//
//                            String key = getKey(text.toString());
//                            String value = mapData.get(key).toString();
//
//                            //nếu thấy \n thì xuống dòng
//                            if(value.contains("\n")){
//                                String [] dataList = value.split("\n");
//                                run.setText(dataList[0], 0);
//                                for (int x = 1; x < dataList.length; x++) {
//                                    String data = dataList[x];
//                                    XWPFRun newRun = paragraph.createRun();
//                                    newRun.setText(data, 0);
//                                    newRun.setFontSize(run.getFontSize());
//                                    newRun.addBreak();
//                                }
//                                break;
//                            } else {
//                                run.setText(value, 0);
//                            }
//                        }
//                    }
//                }
//            }
//
//
//            for (XWPFTable tbl : document.getTables()) {
//                for (XWPFTableRow row : tbl.getRows()) {
//                    for (XWPFTableCell cell : row.getTableCells()) {
//                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
//                            List<XWPFRun> runs = paragraph.getRuns();
//                            if (runs != null) {
//                                for (int i = 0; i < runs.size(); i++) {
//                                    run = runs.get(i);
//                                    if(text != null && !checkSymmetric(text.toString())){
//                                        text.append("\n").append(run.getText(0));
//                                    } else {
//                                        text = new StringBuilder(run.getText(0));
//                                    }
//                                    if (text == null) {
//                                        continue;
//                                    }
//                                    if (checkVariable(text.toString()) || text.toString().contains("$")) {
//                                        while (i < runs.size() - 1) {
//                                            nextRun = runs.get(i + 1);
//                                            text.append(nextRun.getText(0));
//                                            paragraph.removeRun(i + 1);
//                                            if(text.toString().contains("{") && checkSymmetric(text.toString())){
//                                                break;
//                                            }
//                                        }
//                                        if(!checkSymmetric(text.toString())){
////                                            while (paragraph.getRuns().size() > 0) {
////                                                paragraph.removeRun(0);
////                                            }
//                                            continue;
//                                        }
//
//                                        String key = getKey(text.toString());
//                                        String value = mapData.get(key).toString();
//
//                                        //nếu thấy \n thì xuống dòng
//                                        if(value.contains("\n")){
//                                            String [] dataList = value.split("\n");
//                                            for (String data : dataList) {
//                                                XWPFRun newRun = paragraph.createRun();
//                                                newRun.setText(data, 0);
//                                                newRun.setFontSize(run.getFontSize());
//                                                newRun.addBreak();
//                                            }
//                                            paragraph.removeRun(0);
//                                            break;
//                                        } else {
//                                            run.setText(value, 0);
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }
////            -----------------------------------------------------------------
//
//
//
////            List<IBodyElement> bodyElements = document.getBodyElements();
////            for (int i = 0; i < bodyElements.size(); i++) {
////                IBodyElement bodyElement = bodyElements.get(i);
////
////                if (bodyElement instanceof XWPFParagraph) {
////                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
////                    boolean isEmptyParagraph = true;
////                    for (XWPFRun xwpfRun : paragraph.getRuns()) {
////                        if (!xwpfRun.text().trim().isEmpty()) {
////                            isEmptyParagraph = false;
////                            break;
////                        }
////                    }
////                    if (isEmptyParagraph) {
////                        document.removeBodyElement(i);
////                        i--;
////                    }
////                } else if (bodyElement instanceof XWPFTable) {
////                    XWPFTable table = (XWPFTable) bodyElement;
////                    if (table.getRows().size() == 0) {
////                        document.removeBodyElement(i);
////                        i--;
////                    }
////                }
////            }
//
//
//
//
//            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output.docx");
//            document.write(out);
//            out.flush();
//            out.close();
//            inputStream.close();
//            System.out.println("Report generated successfully!");
//            document.close();
//        }
//    }

    //them mới paragraph khi loop data
    public XWPFParagraph insertParagraphAfter(XWPFDocument document, XWPFParagraph currentParagraph) {
        XmlCursor cursor = currentParagraph.getCTP().newCursor();
        cursor.toNextSibling();
        XWPFParagraph nextParagraph = document.insertNewParagraph(cursor);
        nextParagraph.getCTP().setPPr(currentParagraph.getCTP().getPPr());
//        XWPFRun run = nextParagraph.createRun();
//        run.setText(value);
//        run.setFontSize(currentParagraph.getRuns().get(0).getFontSize());

        // Sao chép các run từ paragraph hiện tại sang paragraph mới
        for (XWPFRun run : currentParagraph.getRuns()) {
            XWPFRun newRun = nextParagraph.createRun();
            newRun.getCTR().set(run.getCTR().copy());
        }
        return nextParagraph;
    }


    //replace param trong template
    private void replaceText(XWPFDocument document, String key, String value) throws Exception {
        XWPFRun run, nextRun;
        StringBuilder text;
        for (int p = 0; p < document.getParagraphs().size(); p++) {
            XWPFParagraph paragraph = document.getParagraphs().get(p);
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);
                    text = new StringBuilder(run.getText(0));
                    if (text == null) {
                        continue;
                    }
                    if ((key.contains("List") || key.contains("list")) && text.toString().contains("$")) {
                        StringBuilder textListData = new StringBuilder(text);
                        int x = i;
                        while (x < runs.size() - 1) {
                            x++;
                            nextRun = runs.get(x);
                            textListData.append(nextRun.getText(0));
                            if(checkSymmetric(textListData.toString()) && textListData.toString().contains("{")){
                                break;
                            }
                        }

                        //check nếu param là danh sách
                        if(checkSymmetric(textListData.toString()) && (textListData.toString().contains("List") || textListData.toString().contains("list"))){
                            //check nếu list data ứng với key truyền vào thì mới thực hiện lặp qua data và tạo thêm paragraph mới

                            for (int pm = partyMembers.size() - 1; pm >= 0; pm--){
                                PartyMember partyMember = partyMembers.get(pm);
                                XWPFParagraph xwpfParagraph = insertParagraphAfter(document, paragraph);
                                List<XWPFRun> newRuns = xwpfParagraph.getRuns();
                                Field[] fields = partyMember.getClass().getDeclaredFields();
                                for (int r = 0; r < newRuns.size(); r++){
                                    for (Field field : fields){
                                        if(newRuns.get(r).getText(0).contains("STT")){
                                            String textReplace = newRuns.get(r).getText(0);
                                            String newText = String.valueOf(pm + 1);
                                            newRuns.get(r).setText(textReplace.replace("${STT}", newText), 0);
                                        }
                                        if(newRuns.get(r).getText(0).contains(field.getName())){
                                            String textReplace = newRuns.get(r).getText(0);
                                            field.setAccessible(true);
                                            String newText = field.get(partyMember).toString();
                                            newRuns.get(r).setText(textReplace.replace("${" + field.getName() + "}", newText), 0);
                                        }
                                    }
                                }

                                // xóa bỏ phần key đánh dấu danh sách
                                XWPFRun runNew, nextRunNew;
                                StringBuilder textData = null;
                                Integer runDelete = null;
                                Integer run2Delete = null;
                                for (int r = 0; r < newRuns.size(); r++){
                                    runNew = newRuns.get(r);
                                    if(textData != null && !checkSymmetric(textData.toString())){
                                        textData.append(runNew.getText(0));
                                    } else {
                                        textData = new StringBuilder(runNew.getText(0));
                                    }
                                    if(checkVariable(textData.toString()) && checkSymmetric(textData.toString())){
                                        runDelete = r;
                                    }
                                    if (textData.toString().contains("$") && run.getText(0).contains("$")) {
                                        runDelete = 0;
                                        while (r < newRuns.size() - 1) {
                                            nextRunNew = newRuns.get(r + 1);
                                            String oldText = textData.toString();
                                            textData.append(nextRunNew.getText(0));
                                            if(!checkVariable(textData.toString())){
                                                xwpfParagraph.removeRun(r + 1);
                                            } else if(checkVariable(textData.toString()) && !checkVariable(oldText)){
                                                xwpfParagraph.removeRun(r + 1);
                                            }else {
                                                r++;
                                                if(textData.toString().contains("{") && checkSymmetric(textData.toString())){
                                                    run2Delete = r;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                if(run2Delete != null){
                                    xwpfParagraph.removeRun(run2Delete);
                                    run2Delete = null;
                                }
                                if(runDelete != null){
                                    xwpfParagraph.removeRun(runDelete);
                                    runDelete = null;
                                }
                            }

                            p += partyMembers.size();
                            //remove các run cũ
                            while (!paragraph.getRuns().isEmpty()){
                                paragraph.removeRun(0);
                            }
                        }
                    }
                    if(!runs.isEmpty()) {
                        if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                            while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                nextRun = runs.get(i + 1);
                                text.append(nextRun.getText(0));
                                paragraph.removeRun(i + 1);
                            }

                            //nếu thấy \n thì xuống dòng
                            if (value.contains("\n")) {
                                String[] dataList = value.split("\n");
                                if (text.toString().contains(key)) {
                                    run.setText(dataList[0], 0);
                                    for (int d = 1; d < dataList.length; d++) {
//                                    paragraph = insertParagraphAfter(document, paragraph, dataList[d]);
//                                    XWPFRun newRun = paragraph.createRun();
//                                    newRun.setText(dataList[d], 0);
//                                    newRun.setFontSize(run.getFontSize());
//                                    newRun.addCarriageReturn();
                                    }
                                    p += dataList.length - 1;
                                }
                            } else {
                                run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                            }
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


    //lặp data
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

    //check xem có đối xứng {}
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

    //xóa key khỏi đoạn text cần lặp nhiều data
    public String removeKeyParamList(String key){
        int start = key.indexOf("${", 2);
        int finish = key.lastIndexOf("}");
        return key.substring(start, finish);
    }

    //kiểm tra đây có phải là param dạng $acb{} không
    public boolean checkVariable(String text){
        String regex = "\\$[^{]*\\{";
        Pattern pattern = Pattern.compile(regex);
        boolean b = pattern.matcher(text).find();
        return b;
    }

    //get key từ 1 đoạn text có cả key và value
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
