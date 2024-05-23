package com.example.demoapachepoi.demoExportPDF.service;

import com.example.demoapachepoi.demoExportPDF.DTO.*;
import com.example.demoapachepoi.demoExportPDF.entity.*;
import com.example.demoapachepoi.demoExportPDF.repository.DecisionParamRepository;
import com.example.demoapachepoi.demoExportPDF.utils.GenTextUtil;
import com.example.demoapachepoi.demoExportPDF.utils.SqlQueryUtil;
import lombok.SneakyThrows;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.StringReader;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.dom.Document;
import org.xml.sax.InputSource;

@Service
public class ExportDecisionService {

    @Autowired
    PartyMemberService partyMemberService;

    @Autowired
    DecisionParamRepository decisionParamRepo;

    @Autowired
    SqlQueryUtil sqlQueryUtil;

    @Autowired
    GenTextUtil genTextUtil;

    private final double convertPtToPx = 1.3333343412075;
    private final double convertTwipsToPx = 0.0666666667;

    @SneakyThrows
    public Map<String, List<GenFileDTO>> exportStyle() {
        List<PartyMember> partyMembers = partyMemberService.getUserList();
        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);

            replaceText2(doc, "$partyMemberList{", null, partyMembers, null);
            replaceText2(doc, "${DangBoCapTren}", "ĐẢNG BỘ TẬP ĐOÀN VIETTEL", null, null);
            replaceText2(doc, "${DangUyRQD}", "ĐẢNG ỦY TẬP ĐOÀN VIETTEL", null, null);
            replaceText2(doc, "${TenCBMoi}", "Chi bộ GPS", null, null);
            replaceText2(doc, "${DangCapTrenTrucTiep}", "Đảng bộ tập đoàn NB", null, null);
            replaceText2(doc, "${NhiemKy}", "2020-2022", null, null);
            replaceText2(doc, "${CoQuanSoanThao}", "FPT", null, null);
            replaceText2(doc, "${CacDieuTren}", "1,2,3", null, null); //todo: thêm stt các điều vào đây
            replaceText2(doc, "${SapNhap}", "trên cơ sở sáp nhập Chi bộ Team A và Chi bộ Team B", null, null);
            replaceText2(doc, "${SoLuongDangVien}", "6", null, null);
            replaceText2(doc, "${ChiUyMoi}", "Chi ủy GPS", null, null);
            replaceText2(doc, "${DieuCuoiCung}", "4", null, null);
            replaceText2(doc, "${ChuCaiDau}", "T", null, null);
            replaceText2(doc, "${SoLuongBanPhatHanh}", "7", null, null);
            replaceText2(doc, "${DiaDanh}", "Hà Nội", null, null);
            replaceText2(doc, "${ChucDanhNguoiKy}", "BÍ THƯ", null, null);

            //get param in database
            List<DecisionParams> decisionParams = decisionParamRepo.findAll();
            for (DecisionParams decisionParam : decisionParams) {
                String keyParam = decisionParam.getParamName();
                String sqlQuery = decisionParam.getSqlQuery();

                if(keyParam.toLowerCase().contains("list")){
                    List<Object[]> listData = sqlQueryUtil.getListDataBySqlString(sqlQuery, null);
                    replaceListInParamDynamic(doc, keyParam, listData);
                } else {
                    Map<String, String> paramList = new LinkedHashMap<>();

                    //todo: add cac tham so dieu kien

                    paramList.put("id", "1");
                    String value = sqlQueryUtil.getDataBySqlString(sqlQuery, paramList);
                    replaceText2(doc, keyParam, value, null, null);
                }
            }

            Map<String, String> textMap = returnText(doc);
            List<StyleParagraphDTO> styleMap = returnStyle(doc, partyMembers);

            Map<String, List<GenFileDTO>> stringListMap = mappingDataV2(textMap, styleMap);

            // Save the modified document
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output.docx");
            doc.write(out);
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("Report generated successfully!");
            doc.close();
            return stringListMap;
        }
    }


    @SneakyThrows
    public Map<String, List<GenFileDTO>> exportStyleBTV() {

        List<PartyOrganizationDraft> organizationDrafts = Arrays.asList(
                new PartyOrganizationDraft((long)1, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)2, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)3, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)4, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)1, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)1, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)3, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)1, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)5, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a"),
                new PartyOrganizationDraft((long)3, "AAAA", "To chuc A", (long)2, (long)1, new Date(), "a")
        );

        List<PartyMemberDraft> partyMemberDrafts = Arrays.asList(
                new PartyMemberDraft((long)1, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)2, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)3, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)4, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)2, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)3, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)4, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)5, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)6, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)3, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)2, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)3, (long)1, (long)1,(long)1),
                new PartyMemberDraft((long)5, (long)1, (long)1,(long)1)
        );

        String soLuongBanPhatHanh = genTextUtil.genSoLuongBanPhatHanh(organizationDrafts, partyMemberDrafts);

        ///////////////////////////////////

        List<PartyMember> partyMembers = partyMemberService.getUserList();

        List<TccdMember> tccdList = Arrays.asList(
                new TccdMember("Đảng bộ cơ sở 1", partyMembers)
                ,new TccdMember("Đảng bộ cơ sở 2", partyMembers)
                ,new TccdMember("Chi bộ cơ sở 3", partyMembers)
        );

        try (InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\Thanh-lap-co-BTV.docx")) {
            XWPFDocument doc = new XWPFDocument(inputStream);
            replaceText2(doc, "$ttcdList{", null, null, tccdList);
            replaceText2(doc, "${DangBoCapTren}", "ĐẢNG BỘ TẬP ĐOÀN VIETTEL", null, null);
            replaceText2(doc, "${DangUyRQD}", "ĐẢNG ỦY TẬP ĐOÀN VIETTEL", null, null);
            replaceText2(doc, "${TenCBMoi}", "Chi bộ GPS", null, null);
            replaceText2(doc, "${DangCapTrenTrucTiep}", "Đảng bộ tập đoàn NB", null, null);
            replaceText2(doc, "${NhiemKy}", "2020-2022", null, null);
            replaceText2(doc, "${SoLuongDangCapDuoi}", "3", null, null);
            replaceText2(doc, "${DangUyMoi}", "Đảng ủy GPS", null, null);
            replaceText2(doc, "${SoLuongBCH}", "4", null, null);
            replaceText2(doc, "$dvbchList{", null, partyMembers, null);
            replaceText2(doc, "${SoLuongBTV}", "5", null, null);
            replaceText2(doc, "$dvbtvList{", null, partyMembers, null);
            replaceText2(doc, "${DSLoaiHinhDangCapDuoi}", "Đảng bộ cơ sở", null, null);
            replaceText2(doc, "${DSCapUyCapDuoi}", "Đảng ủy cơ sở", null, null);
            replaceText2(doc, "${CoQuanSoanThao}", "FPT", null, null);
            replaceText2(doc, "${CacDieuTren}", "1,2,3", null, null); //todo: thêm stt các điều vào đây
            replaceText2(doc, "${SapNhap}", "trên cơ sở sáp nhập Chi bộ Team A và Chi bộ Team B", null, null);
            replaceText2(doc, "${DieuCuoiCung}", "4", null, null);
            replaceText2(doc, "${ChuCaiDau}", "T", null, null);
            replaceText2(doc, "${SoLuongBanPhatHanh}", soLuongBanPhatHanh, null, null);
            replaceText2(doc, "${DiaDanh}", "Hà Nội", null, null);
            replaceText2(doc, "${ChucDanhNguoiKy}", "BÍ THƯ", null, null);

            //get param in database
            List<DecisionParams> decisionParams = decisionParamRepo.findAll();
            for (DecisionParams decisionParam : decisionParams) {
                String keyParam = decisionParam.getParamName();
                String sqlQuery = decisionParam.getSqlQuery();

                if(keyParam.toLowerCase().contains("list")){
                    List<Object[]> listData = sqlQueryUtil.getListDataBySqlString(sqlQuery, null);
                    replaceListInParamDynamic(doc, keyParam, listData);
                } else {
                    Map<String, String> paramList = new LinkedHashMap<>();
                    // todo: add cac tham so dieu kien

                    paramList.put("id", "1");
                    String value = sqlQueryUtil.getDataBySqlString(sqlQuery, paramList);
                    replaceText2(doc, keyParam, value, null, null);
                }
            }

            Map<String, String> textMap = returnText(doc);
            List<StyleParagraphDTO> styleMap = returnStyle(doc, partyMembers);
            Map<String, List<GenFileDTO>> stringListMap = mappingDataV2(textMap, styleMap);

            // Save the modified document
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\outputBTV.docx");
            doc.write(out);
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("Report generated successfully!");
            doc.close();
            return stringListMap;
        }
    }

    @SneakyThrows
    public void genFileAfterStep2(Map<String, List<GenFileDTO>> stringListMap) {
        try (
//                InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\template.docx")
                InputStream inputStream = new FileInputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\Thanh-lap-co-BTV.docx")
        ) {
            XWPFDocument doc = new XWPFDocument(inputStream);
            //gen file after step 2
            genFileV2(stringListMap, doc);
            // Save the modified document
//            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\output2.docx");
            FileOutputStream out = new FileOutputStream("D:\\Projects\\GPS\\CTCT\\projects\\demo-apache-poi\\src\\main\\resources\\templates\\outputBTV.docx");
            doc.write(out);
            out.flush();
            out.close();
            inputStream.close();
            System.out.println("File generated successfully!");
            doc.close();
        }
    }

    // mapping text và style
    public Map<String, List<GenFileDTO>> mappingDataV2(Map<String, String> textMap, List<StyleParagraphDTO> styleMap){
        Map<String, List<GenFileDTO>> mapDataMaster = new LinkedHashMap<>();

        textMap.forEach((key, value) -> {
            List<GenFileDTO> genFileDTOS = new ArrayList<>();
            for(int i = 0; i < styleMap.size(); i++){
                StyleParagraphDTO styleParagraphDTO = styleMap.get(i);
                List<StyleRunDTO> valueStyle = styleParagraphDTO.getStyles();
                GenFileDTO genFileDTO = GenFileDTO.builder().build();
                String keyStyle = styleParagraphDTO.getTitle();
                genFileDTO.setTitle(keyStyle);
                genFileDTO.setStyleParagraph(styleParagraphDTO.getStyleParagraph());

                List<GenFileChildrenDTO> genFileChildrenDTOS = new ArrayList<>();
                valueStyle.forEach(entry -> {
                    GenFileChildrenDTO genFileChildrenDTO = GenFileChildrenDTO.builder()
                            .title(entry.getTitle())
                            .styleParagraph(entry.getStyles())
                            .readOnlyList(entry.getReadOnly() != null ? entry.getReadOnly() : new ArrayList<>())
                            .bulletPoint(entry.isBulletPoint())
                            .build();
                    genFileChildrenDTOS.add(genFileChildrenDTO);
                });
                genFileDTO.setChildrens(genFileChildrenDTOS);
                if(value.contains(keyStyle.trim())){
                    StringBuilder text = new StringBuilder();
                    if(!genFileDTOS.isEmpty()){
                        for (int g = 0; g < genFileDTOS.size(); g++){
                            GenFileDTO genFileDTOChild = genFileDTOS.get(g);
                            text.append(genFileDTOChild.getTitle());
                            text.append("\n");
                        }
                        if(value.contains(text.append(genFileDTO.getTitle()))){
                            genFileDTOS.add(genFileDTO);
                        }
                    } else {
                        genFileDTOS.add(genFileDTO);
                    }
                }
            }
            mapDataMaster.put(key, genFileDTOS);
        });
        return mapDataMaster;
    }

    //đọc style từ xml của paragraph
    public String getStyleFromXml(XWPFParagraph paragraph){
        StringBuilder mapStyleParagraph = new StringBuilder();
        try {
            CTPPr xmlString = paragraph.getCTP().getPPr();
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(new InputSource(new StringReader(xmlString.toString())));

            // xem có căn giữa hay không
            NodeList jcList = document.getElementsByTagName("w:jc");
            if (jcList.getLength() > 0) {
                Element jcElement = (Element) jcList.item(0);
                String valAlign = jcElement.getAttribute("w:val");
                mapStyleParagraph.append("text-align:").append(valAlign).append(",");
            }

            // xem thụt đầu dòng bao nhiêu
            NodeList indList = document.getElementsByTagName("w:ind");
            if (indList.getLength() > 0) {
                Element indElement = (Element) indList.item(0);
                String valfirstLine = indElement.getAttribute("w:firstLine");
                if(StringUtils.isNotEmpty(valfirstLine)){
                    //convert twips -> px
                    double valfirstLinePx = Double.parseDouble(valfirstLine) * convertTwipsToPx;
                    mapStyleParagraph.append("valfirstLine:").append(String.format("%.3f",valfirstLinePx)).append("px").append(",");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return mapStyleParagraph.toString();
    }

    //lay bulletPoint or numberingPoint
    public Map<String, String> getBulletPoint(XWPFParagraph paragraph, XWPFDocument doc, Map<String, String> mapNumberingPoint){
        Map<String, String> mapBullet = new HashMap<>();
        StringBuilder styleParagraph  = new StringBuilder();
        String valueNumberPoint;
        StringBuilder endNumberingPoint = new StringBuilder(".");

        String numFmt = paragraph.getNumFmt();
        BigInteger numID = paragraph.getNumID();
        if(numID != null){
            BigInteger ilvl = paragraph.getNumIlvl();
            XWPFNumbering numbering = doc.getNumbering();
            XWPFNum num = numbering.getNum(numID);
            BigInteger abstractNumID = num.getCTNum().getAbstractNumId().getVal();
            XWPFAbstractNum abstractNum = numbering.getAbstractNum(abstractNumID);
            CTAbstractNum ctAbstractNum = abstractNum.getCTAbstractNum();
            CTLvl ctLvl = ctAbstractNum.getLvlArray(ilvl.intValue());

            // Check the bullet type and get value
            if (numFmt.equals("bullet")) {
                String bulletChar = ctLvl.getLvlText().getVal();
                valueNumberPoint = toUnicodeString(bulletChar);
                endNumberingPoint.setLength(0);
            } else {
                // check ket thuc cua numberingPoint la '.' or ')'
                String numberingChar = ctLvl.getLvlText().getVal();
                if(numberingChar.endsWith(")")){
                    endNumberingPoint = new StringBuilder(")");
                }
                // lay gia tri cua numberingPoint
                int numberStart = ctLvl.getStart().getVal().intValue();
                if(mapNumberingPoint.get(abstractNumID.toString()) != null){
                    numberStart = Integer.parseInt(mapNumberingPoint.get(abstractNumID.toString())) + 1;
                }
                mapNumberingPoint.put(abstractNumID.toString(), String.valueOf(numberStart));
                valueNumberPoint = formatNumber(numberStart, numFmt);
            }

            //get style of bullet
            CTRPr rPr = ctLvl.getRPr();
            boolean bold = rPr.isSetB();
            boolean italic = rPr.isSetI();
            boolean underline = rPr.isSetU();
            boolean setStrike = rPr.isSetStrike();
            boolean setVertAlign = rPr.isSetVertAlign();
            boolean setSz = rPr.isSetSz();

            CTPPr pPr = ctLvl.getPPr();
            double distanceBetweenBulletAndText = 0;
            if(pPr != null){
                int valFirstLinePx = pPr.getInd().getHanging().intValue();
                int leftSpace = pPr.getInd().getLeft().intValue();
                distanceBetweenBulletAndText = (double) (leftSpace - valFirstLinePx) / 20;
                styleParagraph.append("valfirstLine:").append(String.format("%.3f", valFirstLinePx * convertTwipsToPx)).append("px").append(",");
            }

            styleParagraph.append("bold:").append(bold).append(",")
                .append("italic:").append(italic).append(",")
                .append("strike:").append(setStrike).append(",")
                ;
            if(underline){
                STUnderline.Enum uEnum = rPr.getU().getVal();
                styleParagraph.append("underline:").append(uEnum).append(",");
            }
            String subscript = STVerticalAlignRun.BASELINE.toString();
            if(setVertAlign){
                STVerticalAlignRun.Enum val = rPr.getVertAlign().getVal();
                subscript = val.toString();
            }
            styleParagraph.append("subscript:").append(subscript.toUpperCase()).append(",");
            if (setSz){
                int fontSize = rPr.getSz().getVal().intValue();
                double fontSizePx = fontSize * convertPtToPx;
                styleParagraph.append("fontSize:").append(String.format("%.3f", fontSizePx)).append("px").append(",");
            } else {
                styleParagraph.append("fontSize:").append(String.format("%.3f", 14 * convertPtToPx)).append("px").append(",");
            }

//            for(int i = 0; i < Math.ceil(distanceBetweenBulletAndText / 14); i++){
//                endNumberingPoint.append(" ");
//            }

            for(int i = 0; i < distanceBetweenBulletAndText / 14; i++){
                endNumberingPoint.append(" ");
            }

            mapBullet.put(valueNumberPoint + endNumberingPoint, styleParagraph.toString());
        }
        return mapBullet;
    }


    // kiem tra xem numberingPoint là dạng format nào
    private static String formatNumber(int number, String format) {
        switch (format) {
            case "upperRoman":
                return toRoman(number).toUpperCase();
            case "lowerRoman":
                return toRoman(number).toLowerCase();
            case "upperLetter":
                return String.valueOf((char) (number + 64)); // A = 65 in ASCII
            case "lowerLetter":
                return String.valueOf((char) (number + 96)); // a = 97 in ASCII
            case "decimal":
            default:
                return String.valueOf(number);
        }
    }

    /**
     *
     * @param number: so dang int
     * @return: tra ve so thap phan
     */
    private static String toRoman(int number) {
        if (number < 1 || number > 3999) return "Invalid Roman Number Value";
        String[] rnThousands = {"", "M", "MM", "MMM"};
        String[] rnHundreds = {"", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"};
        String[] rnTens = {"", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"};
        String[] rnOnes = {"", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"};
        return rnThousands[number / 1000] +
                rnHundreds[(number % 1000) / 100] +
                rnTens[(number % 100) / 10] +
                rnOnes[number % 10];
    }

    // chuyen ky tu dac biet thanh dang unicode
    private static String toUnicodeString(String bulletChar) {
        StringBuilder unicodeString = new StringBuilder();
        for (char c : bulletChar.toCharArray()) {
            unicodeString.append(String.format("\\u%04X", (int) c));
        }
        return unicodeString.toString();
    }

//    bulletPoint styles
    public String styleBulletPoint(String styleFromXml, Map<String, String> bulletPoint, List<StyleRunDTO> styleRunDTOList){
        if(!bulletPoint.isEmpty()){
            Map.Entry<String, String> entry = bulletPoint.entrySet().stream().findFirst().get();
            String keyBullet = entry.getKey();
            String styleBullet = entry.getValue();

            if(styleBullet.contains("valfirstLine")){
                String px = styleBullet.substring(0, styleBullet.indexOf(","));
                if(!styleFromXml.contains(px)){
                    styleFromXml += px += ",";
                }
                styleBullet = styleBullet.replaceFirst(px, "");
            }

            String bullet = StringUtils.stripEnd(keyBullet, " ");
            String spaceBullet = StringUtils.substringAfter(keyBullet, bullet);

            StyleRunDTO bulletDTO = StyleRunDTO.builder()
                    .title(bullet)
                    .styles(styleBullet)
                    .readOnly(new ArrayList<>())
                    .bulletPoint(true)
                    .build();

            StyleRunDTO bulletSpaceDTO = StyleRunDTO.builder()
                    .title(spaceBullet)
                    .styles("")
                    .readOnly(new ArrayList<>())
                    .bulletPoint(false)
                    .build();
            String fontSizeValueStr = StringUtils.substringBetween(styleBullet, "fontSize:", ",");
            if(StringUtils.isNotEmpty(fontSizeValueStr)){
                bulletSpaceDTO.setStyles("fontSize:" + fontSizeValueStr + ",");
            }

            styleRunDTOList.add(bulletDTO);
            styleRunDTOList.add(bulletSpaceDTO);
        }
        return styleFromXml;
    }

    // tra ve style cua van ban gen ra
    /**
     *
     * @param document
     * @param listData: danh sach cac dang vien duoc neu trong quyet dinh
     * @return: tra ve style cua van ban gen ra
     */
    public List<StyleParagraphDTO> returnStyle(XWPFDocument document, List<PartyMember> listData){
        List<StyleParagraphDTO> styleParagraphDTOS = new ArrayList<>();
        XWPFRun run, nextRun;
        StringBuilder text = null;
        // tao 1 map chua cac numbering point, xem stt dang la bao nhieu
        Map<String, String> mapNumberingPoint = new LinkedHashMap<>();

        for(IBodyElement iBodyElement : document.getBodyElements()){
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null && !runs.isEmpty()) {
                    List<StyleRunDTO> styleRunDTOList = new ArrayList<>();
                    String textParagraph = paragraph.getText();

                    // lấy ra style của paragraph
                    String styleFromXml = getStyleFromXml(paragraph);

                    //get bullet
                    Map<String, String> bulletPoint = getBulletPoint(paragraph, document, mapNumberingPoint);
                    //add bullet to first of list run
                    styleFromXml = styleBulletPoint(styleFromXml, bulletPoint, styleRunDTOList);

                    for (int i = 0; i < runs.size(); i++) {
                        run = runs.get(i);
                        if(run.getText(0) == null){
                            continue;
                        }
                        text = new StringBuilder(run.getText(0));
                        while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
                                && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":")
                                && (i < runs.size() - 1 && !runs.get(i + 1).getText(0).startsWith(" -")))
                                && runs.size() > 1 && i < runs.size() - 1
                        ) {
                            nextRun = runs.get(i + 1);
                            text.append(nextRun.getText(0));
                            i++;
                        }
                        //lấy các thuộc tính của run
                        boolean bold = run.isBold();
                        boolean italic = run.isItalic();
                        boolean strike = run.isStrikeThrough();
                        VerticalAlign subscript = run.getSubscript();

                        int fontSize = run.getFontSize();
                        //convert pt to px = pt * 1.3333343412075
                        double fontSizePx = fontSize * convertPtToPx;
                        UnderlinePatterns underline = run.getUnderline();

                        StringBuilder styleBuilder = new StringBuilder();
                        styleBuilder.append("bold:").append(bold).append(",")
                                .append("italic:").append(italic).append(",")
                                .append("fontSize:").append(String.format("%.3f", fontSizePx)).append("px").append(",")
                                .append("underline:").append(underline).append(",")
                                .append("strike:").append(strike).append(",")
                                .append("subscript:").append(subscript).append(",")
                        ;
                        StyleRunDTO styleRunDTO =  StyleRunDTO.builder()
                                .title(text.toString())
                                .styles(styleBuilder.toString())
                                .build();
                        // check text is read only
                        List<String> readOnlyList = new ArrayList<>();
                        for (PartyMember partyMember : listData) {
                            if (text.toString().contains(partyMember.getHoTen())) {
                                if(!readOnlyList.contains(partyMember.getHoTen())){
                                    readOnlyList.add(partyMember.getHoTen());
                                }
                            }
                        }
                        styleRunDTO.setReadOnly(readOnlyList);
                        styleRunDTO.setBulletPoint(false);
                        styleRunDTOList.add(styleRunDTO);
                    }

                    StyleParagraphDTO styleParagraphDTO = StyleParagraphDTO.builder()
                            .title(textParagraph)
                            .styles(styleRunDTOList)
                            .build();
                    if(styleFromXml != null){
                        styleParagraphDTO.setStyleParagraph(styleFromXml);
                    }

                    styleParagraphDTOS.add(styleParagraphDTO);
                }

            } else if (iBodyElement instanceof XWPFTable) {
                XWPFTable tbl = (XWPFTable) iBodyElement;
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {

                            String textParagraph = paragraph.getText();
                            List<StyleRunDTO> styleRunDTOList = new ArrayList<>();
                            // lấy ra style của paragraph
                            String styleFromXml = getStyleFromXml(paragraph);

                            //get style bullet
                            Map<String, String> bulletPoint = getBulletPoint(paragraph, document, mapNumberingPoint);
                            //add bullet to first of list run
                            styleFromXml = styleBulletPoint(styleFromXml, bulletPoint, styleRunDTOList);

                            List<XWPFRun> runs = paragraph.getRuns();
                            if (runs != null && !runs.isEmpty()) {
                                for (int i = 0; i < runs.size(); i++) {
                                    run = runs.get(i);
                                    text = new StringBuilder(run.getText(0));
                                    while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
                                            && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":")
                                            && (i < runs.size() - 1 && !runs.get(i + 1).getText(0).startsWith(" -")))
                                            && runs.size() > 1 && i < runs.size() - 1
                                    ) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        i++;
                                    }
                                    //lấy các thuộc tính của run
                                    boolean bold = run.isBold();
                                    boolean italic = run.isItalic();
                                    boolean strike = run.isStrikeThrough();
                                    VerticalAlign subscript = run.getSubscript();

                                    int fontSize = run.getFontSize();
                                    //convert pt to px = pt * 1.3333343412075
                                    double fontSizePx = fontSize * convertPtToPx;
                                    UnderlinePatterns underline = run.getUnderline();

                                    StringBuilder styleBuilder = new StringBuilder();
                                    styleBuilder.append("bold:").append(bold).append(",")
                                            .append("italic:").append(italic).append(",")
                                            .append("fontSize:").append(String.format("%.3f", fontSizePx)).append(",")
                                            .append("underline:").append(underline).append(",")
                                            .append("strike:").append(strike).append(",")
                                            .append("subscript:").append(subscript).append(",")
                                    ;
                                    StyleRunDTO styleRunDTO =  StyleRunDTO.builder()
                                            .title(text.toString())
                                            .styles(styleBuilder.toString())
                                            .build();

                                    // check text is read only
                                    List<String> readOnlyList = new ArrayList<>();
                                    for (PartyMember partyMember : listData) {
                                        if (text.toString().contains(partyMember.getHoTen()) && (!readOnlyList.contains(partyMember.getHoTen()))){
                                            readOnlyList.add(partyMember.getHoTen());
                                        }
                                    }
                                    styleRunDTOList.add(styleRunDTO);
                                }
                                StyleParagraphDTO styleParagraphDTO = StyleParagraphDTO.builder()
                                        .title(textParagraph)
                                        .styles(styleRunDTOList)
                                        .build();
                                if(styleFromXml != null){
                                    styleParagraphDTO.setStyleParagraph(styleFromXml);
                                }
                                styleParagraphDTOS.add(styleParagraphDTO);
                            }
                        }
                    }
                }
            }
        }
        return styleParagraphDTOS;
    }


    //test
    public Map<String, List<GenFileDTO>> returnStyleV2(XWPFDocument document){
        Map<String, List<GenFileDTO>> textGenFile = new LinkedHashMap<>();
        List<GenFileDTO> genFileDTOS = new ArrayList<>();
        // tao 1 map chua cac numbering point, xem stt dang la bao nhieu
        Map<String, String> mapNumberingPoint = new LinkedHashMap<>();

        XWPFRun run, nextRun;
        StringBuilder text = null;
        StringBuilder textMultiParagraph = null;

        for(IBodyElement iBodyElement : document.getBodyElements()){
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                List<XWPFRun> runs = paragraph.getRuns();

                if (runs != null && !runs.isEmpty()) {
                    GenFileDTO genFileDTO = new GenFileDTO();

                    //set title
                    String textParagraph = paragraph.getText();
                    genFileDTO.setTitle(textParagraph);

                    // get style of paragraph
                    String styleParagraphFromXml = getStyleFromXml(paragraph);


                    //set style of paragraph
                    if(styleParagraphFromXml != null){
                        genFileDTO.setStyleParagraph(styleParagraphFromXml);
                    }

                    //get style all run
                    List<GenFileChildrenDTO> genFileChildrenDTOS = new ArrayList<>();
                    for (int i = 0; i < runs.size(); i++) {
                        run = runs.get(i);
                        if(run.getText(0) == null){
                            continue;
                        }
                        text = new StringBuilder(run.getText(0));
                        int runIndex = i;

                        if((!checkVariable(text.toString()) && !text.toString().contains("}")) ||
                            (text.toString().contains("}") && textMultiParagraph != null &&
                                checkVariable(textMultiParagraph.toString()) && !checkSymmetric(textMultiParagraph.append(text).toString())
                                && i == runs.size() -1
                            )
                        ){
                            while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
                                    && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":")
                                    && (runIndex < runs.size() - 1 && !runs.get(runIndex + 1).getText(0).startsWith(" -"))
                                    && !runs.get(runIndex + 1).getText(0).startsWith("}") )
                                    && runs.size() > 1 && runIndex < runs.size() - 1
                            ) {
                                nextRun = runs.get(runIndex + 1);
                                text.append(nextRun.getText(0));
                                runIndex++;
                            }
                            //lấy các thuộc tính của run
                            boolean bold = run.isBold();
                            boolean italic = run.isItalic();
                            String color = run.getColor();

                            int fontSize = run.getFontSize();
                            //convert pt to px = pt * 1.3333343412075
                            double fontSizePx = fontSize * convertPtToPx;
                            UnderlinePatterns underline = run.getUnderline();

                            StringBuilder styleBuilder = new StringBuilder();
                            styleBuilder.append("bold:").append(bold).append(",")
                                    .append("italic:").append(italic).append(",")
                                    .append("color:").append(color).append(",")
                                    .append("fontSize:").append(String.format("%.3f", fontSizePx)).append(",")
                                    .append("underline:").append(underline).append(",")
                            ;
                            GenFileChildrenDTO genFileChildrenDTO = GenFileChildrenDTO.builder()
                                    .title(text.toString())
                                    .styleParagraph(styleBuilder.toString())
                                    .build();
                            genFileChildrenDTOS.add(genFileChildrenDTO);
                        }
                        i = runIndex;
                    }
                    genFileDTO.setChildrens(genFileChildrenDTOS);
                    genFileDTOS.add(genFileDTO);

                    //check is multi paragraph
                    for (int i = 0; i < runs.size(); i++) {
                        run = runs.get(i);
                        if (run.getText(0) == null) {
                            continue;
                        }

                        if(textMultiParagraph != null && !checkSymmetric(textMultiParagraph.toString())){
                            textMultiParagraph.append("\n").append(run.getText(0));
                        } else {
                            textMultiParagraph = new StringBuilder(run.getText(0));
                        }
                        if (checkVariable(textMultiParagraph.toString()) || textMultiParagraph.toString().contains("$")) {
                            int pa = i;
                            while (pa < runs.size() - 1) {
                                nextRun = runs.get(pa + 1);
                                textMultiParagraph.append(nextRun.getText(0) != null ? nextRun.getText(0) : "");
                                pa++;
                                if(textMultiParagraph.toString().contains("{") && checkSymmetric(textMultiParagraph.toString())){
                                    break;
                                }
                            }
                            if(checkSymmetric(textMultiParagraph.toString())){
                                textGenFile.put(getKey(textMultiParagraph.toString()), genFileDTOS);
                                genFileDTOS.clear();
                                break;
                            }
                        }
                    }
                }

            }
//            else if (iBodyElement instanceof XWPFTable) {
//                XWPFTable tbl = (XWPFTable) iBodyElement;
//                for (XWPFTableRow row : tbl.getRows()) {
//                    for (XWPFTableCell cell : row.getTableCells()) {
//                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
//
//                            String textParagraph = paragraph.getText();
//                            Map<String, String> map = new LinkedHashMap<>();
//                            // lấy ra style của paragraph
//                            CTPPr ppr = paragraph.getCTP().getPPr();
//                            String mapStyleFromXml = getStyleFromXml(ppr.toString());
//                            if(mapStyleFromXml != null){
//                                map.put("styleParagraph", mapStyleFromXml);
//                            }
//                            List<XWPFRun> runs = paragraph.getRuns();
//                            if (runs != null && !runs.isEmpty()) {
//                                for (int i = 0; i < runs.size(); i++) {
//                                    run = runs.get(i);
//                                    text = new StringBuilder(run.getText(0));
//                                    while ((!text.toString().endsWith(" ") && !text.toString().endsWith(".") && !text.toString().endsWith(";")
//                                            && !text.toString().endsWith("!") && !text.toString().endsWith(",") && !text.toString().endsWith(":")
//                                            && (i < runs.size() - 1 && !runs.get(i + 1).getText(0).startsWith(" -")))
//                                            && runs.size() > 1 && i < runs.size() - 1
//                                    ) {
//                                        nextRun = runs.get(i + 1);
//                                        text.append(nextRun.getText(0));
//                                        i++;
//                                    }
//                                    //lấy các thuộc tính của run
//                                    boolean bold = run.isBold();
//                                    boolean italic = run.isItalic();
//                                    String color = run.getColor();
//
//                                    int fontSize = run.getFontSize();
//                                    //convert pt to px = pt * 1.3333343412075
//                                    double fontSizePx = fontSize * convertPtToPx;
//                                    UnderlinePatterns underline = run.getUnderline();
//
//                                    StringBuilder styleBuilder = new StringBuilder();
//                                    styleBuilder.append("bold:").append(bold).append(",")
//                                            .append("italic:").append(italic).append(",")
//                                            .append("color:").append(color).append(",")
//                                            .append("fontSize:").append(String.format("%.3f", fontSizePx)).append(",")
//                                            .append("underline:").append(underline).append(",")
//                                    ;
//                                    map.put(text.toString(), styleBuilder.toString());
//                                }
//                                mapStyle2.put(textParagraph, map);
//                            }
//                        }
//                    }
//                }
//            }
        }
        return textGenFile;
    }


    // tra ve text cho step2 tu van ban gen ra
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
                        if(run.getText(0) == null){
                            continue;
                        }
                        if(text != null && !checkSymmetric(text.toString())){
                            text.append("\n").append(run.getText(0));
                        } else {
                            text = new StringBuilder(run.getText(0));
                        }
                        if (checkVariable(text.toString()) || text.toString().contains("$")) {
                            while (i < runs.size() - 1) {
                                nextRun = runs.get(i + 1);
                                text.append(nextRun.getText(0) != null ? nextRun.getText(0) : "");
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
                                            text.append(nextRun.getText(0) != null ? nextRun.getText(0) : "");
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
        deleteKey(document);
        return mapData;
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
                for (XWPFTableCell cell : row.getTableCells()) {
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


    //delete key cua từng paragraph đã được xác định
    public void deleteKeyPara(XWPFParagraph xwpfParagraph){
        List<XWPFRun> newRuns = xwpfParagraph.getRuns();
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
            if (textData.toString().contains("$") && runNew.getText(0).contains("$")) {
                runDelete = 0;
                if(runNew.getText(0).contains("$") &&  runDelete != null && runDelete == 0){
                    runDelete = r;
                }
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

    //xoa bo key danh dau danh sach thanh vien trong to chuc(danh sach long nhau)
    public void deleteKeyParaList(XWPFParagraph xwpfParagraphParent, XWPFParagraph xwpfParagraphChild){
        String textParagraph = xwpfParagraphParent.getParagraphText() + xwpfParagraphChild.getParagraphText();
        if(xwpfParagraphChild.getRuns().size() > 0){
            if(checkSymmetric(textParagraph)){
                XWPFRun lastRun = xwpfParagraphChild.getRuns().get(xwpfParagraphChild.getRuns().size() - 1);
                String textLastRun = lastRun.getText(0);
                XWPFRun lastRun2 = xwpfParagraphChild.getRuns().get(xwpfParagraphChild.getRuns().size() - 2);
                if(textLastRun.contains("}}")){
                    lastRun.setText("}", 0);
                } else if(textLastRun.contains("}")){
                    lastRun.setText("", 0);
                } else {
                    lastRun2.setText("", 0);
                }
            }
        }

        List<XWPFRun> newRuns = xwpfParagraphChild.getRuns();
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
            if (textData.toString().contains("$") && runNew.getText(0).contains("$")) {
                runDelete = 0;
                while (r < newRuns.size() - 1) {
                    nextRunNew = newRuns.get(r + 1);
                    String oldText = textData.toString();
                    textData.append(nextRunNew.getText(0));
                    if(!checkVariable(textData.toString())){
                        xwpfParagraphChild.removeRun(r + 1);
                    } else if(checkVariable(textData.toString()) && !checkVariable(oldText)){
                        xwpfParagraphChild.removeRun(r + 1);
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
            xwpfParagraphChild.removeRun(run2Delete);
            run2Delete = null;
        }
        if(runDelete != null){
            xwpfParagraphChild.removeRun(runDelete);
            runDelete = null;
        }
    }


    //them moi cac paragraph vao sau paragraph hien tai
    public XWPFParagraph insertParaAfterCurrentPara(XWPFDocument document, XWPFParagraph currentParagraph) {
        XmlCursor cursor = currentParagraph.getCTP().newCursor();
        cursor.toNextSibling();
        XWPFParagraph nextParagraph = document.insertNewParagraph(cursor);
        nextParagraph.getCTP().setPPr(currentParagraph.getCTP().getPPr());

        // Sao chép các run từ paragraph hiện tại sang paragraph mới
        for (XWPFRun run : currentParagraph.getRuns()) {
            XWPFRun newRun = nextParagraph.createRun();
            newRun.getCTR().set(run.getCTR().copy());
        }
        return nextParagraph;
    }


    //them moi cac paragraph case danh sach long nhau
    public XWPFParagraph insertParaListAfterCurrentPara(XWPFDocument document, XWPFParagraph currentParagraphPosition, XWPFParagraph currentParagraph) {
        XmlCursor cursor;
        if(currentParagraphPosition != null){
            cursor = currentParagraphPosition.getCTP().newCursor();
        } else {
            cursor = currentParagraph.getCTP().newCursor();
        }
        cursor.toNextSibling();
        XWPFParagraph nextParagraph = document.insertNewParagraph(cursor);
        nextParagraph.getCTP().setPPr(currentParagraph.getCTP().getPPr());

        // Sao chép các run từ paragraph hiện tại sang paragraph mới
        for (XWPFRun run : currentParagraph.getRuns()) {
            XWPFRun newRun = nextParagraph.createRun();
            newRun.getCTR().set(run.getCTR().copy());
        }
        return nextParagraph;
    }

    //them para sau khi xoa cac para cu sau buoc 2
    public XWPFParagraph addParagraph(XWPFDocument document, XWPFParagraph currentParagraph){
        XmlCursor cursor = currentParagraph.getCTP().newCursor();
        cursor.toNextSibling();
        XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
        return newParagraph;
    }


    //replace param trong template // su dung cho case van ban co danh sach khong co danh sach con // khong dung nua
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
//                            replaceList(document, paragraph, partyMembers);
//                            p += partyMembers.size();
                            //remove các run cũ
                            while (!paragraph.getRuns().isEmpty()){
                                paragraph.removeRun(0);
                            }
                        }
                    }
                    // cac case khong phai la danh sach
                    if(!runs.isEmpty()) {
                        if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))
                                && text.toString().concat(runs.get(i + 1).getText(0)) .contains("${")
                        ) {
                            while (i < runs.size() - 1 && !text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                nextRun = runs.get(i + 1);
                                text.append(nextRun.getText(0));
                                paragraph.removeRun(i + 1);
                            }
                            //doi param thanh text
                            run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                        }
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

    //replace danh sach khong co danh sach con// khong dung nua
    @SneakyThrows
    public void replaceList(XWPFDocument document, XWPFParagraph paragraph, List<PartyMember> partyMembers){
        for (int pm = partyMembers.size() - 1; pm >= 0; pm--){
            PartyMember partyMember = partyMembers.get(pm);
            XWPFParagraph xwpfParagraph = insertParaAfterCurrentPara(document, paragraph);
            List<XWPFRun> newRuns = xwpfParagraph.getRuns();
            Field[] fields = partyMember.getClass().getDeclaredFields();
            for (int r = 0; r < newRuns.size(); r++){
                for (Field field : fields){
                    String textReplace = newRuns.get(r).getText(0);
                    if(textReplace.contains("STT")){
                        String newText = String.valueOf(pm + 1);
                        newRuns.get(r).setText(textReplace.replace("${STT}", newText), 0);
                    }
                    if(textReplace.contains(field.getName())){
                        field.setAccessible(true);
                        String newText = field.get(partyMember).toString();
                        newRuns.get(r).setText(textReplace.replace("${" + field.getName() + "}", newText), 0);
                    }
                }
            }
            // xóa bỏ phần key đánh dấu danh sách
//            deleteKeyPara(xwpfParagraph);
        }
    }

    // replace danh sach khong co danh sach con v2
    @SneakyThrows
    public void replaceListV2(XWPFDocument document, XWPFParagraph paragraph, List<PartyMember> partyMembers, List<DecisionParams> decisionParams){

        for (int pm = partyMembers.size() - 1; pm >= 0; pm--){
            PartyMember partyMember = partyMembers.get(pm);
            XWPFParagraph xwpfParagraph = insertParaAfterCurrentPara(document, paragraph);
            List<XWPFRun> newRuns = xwpfParagraph.getRuns();
            Field[] fields = partyMember.getClass().getDeclaredFields();
            for (Field field : fields){
                StringBuilder textRunChild = null;
                for (int r = 0; r < newRuns.size(); r++){
                    XWPFRun runCurrent = newRuns.get(r);
                    if(textRunChild != null && !checkSymmetric(textRunChild.toString())){
                        textRunChild.append(runCurrent.getText(0));
                        if(textRunChild.toString().contains("${") && checkSymmetric(textRunChild.toString())){
                            newRuns.get(r - 1).setText(textRunChild.toString(), 0);
                            xwpfParagraph.removeRun(r);
                            r--;
                        } else {
                            xwpfParagraph.removeRun(r);
                            r--;
                            continue;
                        }
                    } else {
                        textRunChild = new StringBuilder(runCurrent.getText(0));
                    }

                    if(r < newRuns.size() - 1 && textRunChild.toString().contains("${")
                            && !checkSymmetric(textRunChild.toString())
                    ){
                        String nextRunTextChild = newRuns.get(r + 1).getText(0);
                        textRunChild.append(nextRunTextChild);
                        runCurrent.setText(textRunChild.toString(), 0);
                        xwpfParagraph.removeRun(r + 1);
                        continue;
                    }
                    textRunChild = null;
                    String textReplace = runCurrent.getText(0);
                    if(textReplace.contains("STT")){
                        String stt = String.valueOf(pm + 1);
                        runCurrent.setText(textReplace.replace("${STT}", stt), 0);
                    }
                    if(textReplace.contains(field.getName())){
                        field.setAccessible(true);
                        String newText = field.get(partyMember).toString();
                        runCurrent.setText(textReplace.replace("${" + field.getName() + "}", newText), 0);
                    } else {
                        // lap qua cac param trong db, neu tim thay -> replace = data query tu cau sql
                        for(DecisionParams decisionParam : decisionParams){
                            String key = decisionParam.getParamName();
                            String sqlQuery = decisionParam.getSqlQuery();
                            if(key.toLowerCase().contains("list")){
                                continue;
                            }

                            // todo:add cac tham so dieu kien
                            Map<String, String> paramList = new LinkedHashMap<>();
                            paramList.put("id", "1");

                            String value = sqlQueryUtil.getDataBySqlString(sqlQuery, paramList);
                            if(textReplace.contains(key)){
                                runCurrent.setText(textReplace.replace(key, value), 0);
                            }
                        }
                    }
                }
            }
            // xóa bỏ phần key đánh dấu danh sách
            deleteKeyPara(xwpfParagraph);
        }
    }

    // dung cho case replace list param trong db
    @SneakyThrows
    public void replaceListParamDynamic(XWPFDocument document, XWPFParagraph paragraph, List<Object[]> objectsList){

        for (int pm = objectsList.size() - 1; pm >= 0; pm--){
            Object[] data = objectsList.get(pm);
            int indexData = 0;

            XWPFParagraph xwpfParagraph = insertParaAfterCurrentPara(document, paragraph);
            List<XWPFRun> newRuns = xwpfParagraph.getRuns();

            StringBuilder textRunChild = null;
            for (int r = 0; r < newRuns.size(); r++){
                XWPFRun runCurrent = newRuns.get(r);
                if(textRunChild != null && !checkSymmetric(textRunChild.toString())){
                    textRunChild.append(runCurrent.getText(0));
                    if(textRunChild.toString().contains("${") && checkSymmetric(textRunChild.toString())){
                        newRuns.get(r - 1).setText(textRunChild.toString(), 0);
                        xwpfParagraph.removeRun(r);
                        r--;
                    } else {
                        xwpfParagraph.removeRun(r);
                        r--;
                        continue;
                    }
                } else {
                    textRunChild = new StringBuilder(runCurrent.getText(0));
                }

                if(r < newRuns.size() - 1 && textRunChild.toString().contains("${")
                        && !checkSymmetric(textRunChild.toString())
                ){
                    String nextRunTextChild = newRuns.get(r + 1).getText(0);
                    textRunChild.append(nextRunTextChild);
                    runCurrent.setText(textRunChild.toString(), 0);
                    xwpfParagraph.removeRun(r + 1);
                    continue;
                }
                textRunChild = null;
                String textReplace = runCurrent.getText(0);
                if(textReplace.contains("STT")){
                    String stt = String.valueOf(pm + 1);
                    runCurrent.setText(textReplace.replace("${STT}", stt), 0);
                }
                String [] currentTextSplit = runCurrent.getText(0).split(" ");
                for (int ct = 0; ct < currentTextSplit.length; ct++){
                    String currentText = currentTextSplit[ct];
                    if(checkSymmetric(currentText) && currentText.startsWith("${")){
                        currentTextSplit[ct] = data[indexData].toString();
                        indexData++;
                    }
                }
                runCurrent.setText(String.join(" ", currentTextSplit).concat(" "), 0);
            }
            // xóa bỏ phần key đánh dấu danh sách
            deleteKeyPara(xwpfParagraph);
        }
    }

    //replace danh sach lồng nhau
    @SneakyThrows
    public void replaceNestedList(XWPFDocument document, XWPFParagraph paragraphParentOld, XWPFParagraph paragraphChildOld,
                                  List<TccdMember> tccdList, List<DecisionParams> decisionParams
    ){
        XWPFParagraph paragraphParentNew = null;
        for (int tc = tccdList.size() - 1; tc >= 0; tc--){
            TccdMember tccdMember = tccdList.get(tc);
            //tao ra paragraph cha(tuong ung to chuc cha)
            XWPFParagraph parParent = insertParaListAfterCurrentPara(document, null, paragraphParentOld);
            paragraphParentNew = parParent;

            //lap qua cac run de replace text
            List<XWPFRun> newRunsParent = parParent.getRuns();
            Field[] fieldsParent = tccdMember.getClass().getDeclaredFields();
            for (Field field : fieldsParent){
                StringBuilder textRunChild = null;
                for (int r = 0; r < newRunsParent.size(); r++){
                    if(textRunChild != null && !checkSymmetric(textRunChild.toString())){
                        textRunChild.append(newRunsParent.get(r).getText(0));
                        if(textRunChild.toString().contains("${") && checkSymmetric(textRunChild.toString())){
                            newRunsParent.get(r - 1).setText(textRunChild.toString(), 0);
                            parParent.removeRun(r);
                            r--;
                        }
                    } else {
                        textRunChild = new StringBuilder(newRunsParent.get(r).getText(0));
                    }

                    if(r < newRunsParent.size() - 1 && textRunChild.toString().contains("${")
                            && !checkSymmetric(textRunChild.toString())
                    ){
                        String nextRunTextChild = newRunsParent.get(r + 1).getText(0);
                        textRunChild.append(nextRunTextChild);
                        newRunsParent.get(r).setText(textRunChild.toString(), 0);
                        parParent.removeRun(r + 1);
                        continue;
                    }
                    textRunChild = null;
                    String textReplace = newRunsParent.get(r).getText(0);
                    if(textReplace.contains("STT")){
                        String stt = String.valueOf(tc + 1);
                        newRunsParent.get(r).setText(textReplace.replace("${STT}", stt), 0);
                    }
                    if(textReplace.contains(field.getName())){
                        field.setAccessible(true);
                        String newText = field.get(tccdMember).toString();
                        newRunsParent.get(r).setText(textReplace.replace("${" + field.getName() + "}", newText), 0);
                    }
                    else {
                        // lap qua cac param trong db, neu tim thay -> replace = data query tu cau sql
                        for(DecisionParams decisionParam : decisionParams){
                            String key = decisionParam.getParamName();
                            String sqlQuery = decisionParam.getSqlQuery();
                            if(key.toLowerCase().contains("list")){
                                continue;
                            }

                            // todo:add cac tham so dieu kien
                            Map<String, String> paramList = new LinkedHashMap<>();
                            paramList.put("id", "1");

                            String value = sqlQueryUtil.getDataBySqlString(sqlQuery, paramList);
                            if(textReplace.contains(key)){
                                newRunsParent.get(r).setText(textReplace.replace(key, value), 0);
                            }
                        }
                    }
                }
            }

            // lap qua cac thanh vien trong to chuc con
            for (int pm = tccdMember.getPartyMembers().size() - 1; pm >= 0; pm--){
                PartyMember partyMember = tccdMember.getPartyMembers().get(pm);

                // tao ra cac paragraph con
                XWPFParagraph xwpfParagraph = insertParaListAfterCurrentPara(document, paragraphParentNew, paragraphChildOld);
                List<XWPFRun> newRuns = xwpfParagraph.getRuns();
                Field[] fields = partyMember.getClass().getDeclaredFields();
                for (Field field : fields){
                    StringBuilder textRunChild = null;
                    for (int r = 0; r < newRuns.size(); r++){
                        if(textRunChild != null && !checkSymmetric(textRunChild.toString())){
                            textRunChild.append(newRuns.get(r).getText(0));
                            if(textRunChild.toString().contains("${") && checkSymmetric(textRunChild.toString())){
                                newRuns.get(r - 1).setText(textRunChild.toString(), 0);
                                xwpfParagraph.removeRun(r);
                                r--;
                            } else {
                                xwpfParagraph.removeRun(r);
                                r--;
                                continue;
                            }
                        } else {
                            textRunChild = new StringBuilder(newRuns.get(r).getText(0));
                        }

                        if(r < newRuns.size() - 1 && textRunChild.toString().contains("${")
                                && !checkSymmetric(textRunChild.toString())
                        ){
                            String nextRunTextChild = newRuns.get(r + 1).getText(0);
                            textRunChild.append(nextRunTextChild);
                            newRuns.get(r).setText(textRunChild.toString(), 0);
                            xwpfParagraph.removeRun(r + 1);
                            continue;
                        }
                        textRunChild = null;
                        String textReplace = newRuns.get(r).getText(0);
                        if(textReplace.contains("STT")){
                            String stt = String.valueOf(pm + 1);
                            newRuns.get(r).setText(textReplace.replace("${STT}", stt), 0);
                        }
                        if(textReplace.contains(field.getName())){
                            field.setAccessible(true);
                            String newText = field.get(partyMember).toString();
                            newRuns.get(r).setText(textReplace.replace("${" + field.getName() + "}", newText), 0);
                        }
                        else {
                            // lap qua cac param trong db, neu tim thay -> replace = data query tu cau sql
                            for(DecisionParams decisionParam : decisionParams){
                                String key = decisionParam.getParamName();
                                String sqlQuery = decisionParam.getSqlQuery();
                                if(key.toLowerCase().contains("list")){
                                    continue;
                                }

                                // todo:add cac tham so dieu kien
                                Map<String, String> paramList = new LinkedHashMap<>();
                                paramList.put("id", "1");

                                String value = sqlQueryUtil.getDataBySqlString(sqlQuery, paramList);
                                if(textReplace.contains(key)){
                                    newRuns.get(r).setText(textReplace.replace(key, value), 0);
                                }
                            }
                        }
                    }
                }
                // xóa bỏ phần key đánh dấu danh sách thanh vien
                deleteKeyParaList(parParent, xwpfParagraph);
            }
            //xoa bo phan key danh dau danh sach to chuc con
            deleteKeyPara(parParent);
        }
    }


//    replace text v2
    private void replaceText2(XWPFDocument document, String key, String value, List<PartyMember> partyMembers, List<TccdMember> tccdList) throws Exception {

        // lay cac param trong db theo id quyet dinh
        List<DecisionParams> decisionParams = decisionParamRepo.findAll();

        XWPFRun run, nextRun;
        StringBuilder text = null;
        boolean appendData = false; // dung de check xem neu para chua phai la 1 doan -> se append data para truoc vao text
        int posOfParaParent = 0; //xac dinh vi tri cua paragraph to chuc con

        for (int p = 0; p < document.getParagraphs().size(); p++) {
            XWPFParagraph paragraph = document.getParagraphs().get(p);
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);

                    if(StringUtils.isNotEmpty(run.getText(0))){
                        if(text != null && !checkSymmetric(text.toString()) && appendData){
                            text.append(run.getText(0));
                        } else {
                            text = new StringBuilder(run.getText(0));
                        }
                        
                        if (text == null) {
                            continue;
                        }

                        //check neu param la danh sach
                        if (key.toLowerCase().contains("list") && text.toString().startsWith("$")
                            && ( i < runs.size() - 1 &&  runs.get(i + 1).getText(0) != null && text.toString().concat(runs.get(i + 1).getText(0)).toLowerCase().contains("list"))
                        ) {
                            int x = i;
                            while (x < runs.size() - 1) {
                                x++;
                                nextRun = runs.get(x);
                                text.append(nextRun.getText(0));
                                if(checkSymmetric(text.toString()) && text.toString().contains("{")){
                                    break;
                                }
                            }

                            //check neu key == key in text, else -> continue
                            if(!StringUtils.startsWith(text.toString(), key)){
                              continue;
                            }

                            // case danh sach khong long nhau
                            if(checkSymmetric(text.toString()) && !appendData){
                                if(partyMembers != null && !partyMembers.isEmpty()){
                                    replaceListV2(document, paragraph, partyMembers, decisionParams);
                                    p += partyMembers.size();
                                    //remove các run cũ
                                    while (!paragraph.getRuns().isEmpty()){
                                        paragraph.removeRun(0);
                                    }
                                }
                            } else { //case danh sach long nhau
                                int posOfParaChild;
                                if(checkVariable(text.toString()) && !checkSymmetric(text.toString())){
                                    appendData = true;
                                    // lay ra vi tri cua paragraph cha
                                    posOfParaParent = document.getPosOfParagraph(paragraph);
                                    break;
                                } else {
                                    //lay ra vi tri cua paragraph con
                                    posOfParaChild = document.getPosOfParagraph(paragraph);
                                    appendData = false;
                                }

                                //check nếu param là danh sách
                                if(checkSymmetric(text.toString()) && text.toString().toLowerCase().contains("list")
                                ){
                                    if(tccdList != null && !tccdList.isEmpty()) {
                                        // lay ra paragraph cha(tuong ung to chuc)
                                        XWPFParagraph paragraphParentOld = document.getParagraphs().get(posOfParaParent - 1);
                                        // lay ra paragraph con(tuong ung thanh vien)
                                        XWPFParagraph paragraphChildOld = document.getParagraphs().get(posOfParaChild - 1);

                                        //check nếu list data ứng với key truyền vào thì mới thực hiện lặp qua data và tạo thêm paragraph mới
                                        replaceNestedList(document, paragraphParentOld, paragraphChildOld, tccdList, decisionParams);
                                        p += tccdList.size();

                                        //remove các run cũ cua para to chuc con va para cac thanh vien
                                        while (!paragraphParentOld.getRuns().isEmpty()) {
                                            paragraphParentOld.removeRun(0);
                                        }
                                        while (!paragraphChildOld.getRuns().isEmpty()) {
                                            paragraphChildOld.removeRun(0);
                                        }
                                    }
                                }
                            }
                        }
                        // cac case khong phai la danh sach
                        if(!runs.isEmpty()) {
                            if(value != null){
                                if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))
                                    && text.toString().concat(runs.get(i + 1).getText(0)) .contains("${")
                                ) {
                                    while (i < runs.size() - 1 && !text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        paragraph.removeRun(i + 1);
                                    }
                                    //doi param thanh text
                                    run.setText(text.toString().contains(key) ? text.toString().replace(key, value) : text.toString(), 0);
                                }
                            }
                        }
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
                                text = new StringBuilder(run.getText(0));
                                if (text == null) {
                                    continue;
                                }
                                if(value == null) break;
                                if (text.toString().contains("${") || (text.toString().contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))
                                        && text.toString().concat(runs.get(i + 1).getText(0)) .contains("${")
                                ) {
                                    while (!text.toString().contains("}") || !checkSymmetric(text.toString())) {
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        paragraph.removeRun(i + 1);
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


    // dung cho cac case list data trong cac param cau hinh trong DB
    public void replaceListInParamDynamic(XWPFDocument document, String key, List<Object[]> listData){

        XWPFRun run, nextRun;
        StringBuilder text = null;
        boolean appendData = false; // dung de check xem neu para chua phai la 1 doan -> se append data para truoc vao text

        for (int p = 0; p < document.getParagraphs().size(); p++) {
            XWPFParagraph paragraph = document.getParagraphs().get(p);
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (int i = 0; i < runs.size(); i++) {
                    run = runs.get(i);

                    if(StringUtils.isNotEmpty(run.getText(0))){
                        if(text != null && !checkSymmetric(text.toString()) && text.toString().contains("$")){
                            text.append(run.getText(0));
                        } else {
                            text = new StringBuilder(run.getText(0));
                        }

                        if (text == null) {
                            continue;
                        }

                        //check neu param la danh sach
                        if (key.toLowerCase().contains("list") && text.toString().startsWith("$")
                                && ( i < runs.size() - 1 &&  runs.get(i + 1).getText(0) != null && text.toString().concat(runs.get(i + 1).getText(0)).toLowerCase().contains("list"))
                        ) {
                            int x = i;
                            while (x < runs.size() - 1) {
                                x++;
                                nextRun = runs.get(x);
                                text.append(nextRun.getText(0));
                                if(checkSymmetric(text.toString()) && text.toString().contains("{")){
                                    break;
                                }
                            }

                            //check neu key == key in text, else -> continue
                            if(StringUtils.startsWith(text.toString(), key)){
                                if(checkSymmetric(text.toString()) && !appendData){
                                    if(listData != null && !listData.isEmpty()){
                                        replaceListParamDynamic(document, paragraph, listData);
                                        p += listData.size();
                                        //remove các run cũ
                                        while (!paragraph.getRuns().isEmpty()){
                                            paragraph.removeRun(0);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (int p = 0; p < cell.getParagraphs().size(); p++) {
                        XWPFParagraph paragraph = cell.getParagraphs().get(p);
                        List<XWPFRun> runs = paragraph.getRuns();
                        if (runs != null) {
                            for (int i = 0; i < runs.size(); i++) {
                                run = runs.get(i);
                                if(StringUtils.isNotEmpty(run.getText(0))){
                                    if(text != null && !checkSymmetric(text.toString())){
                                        text.append(run.getText(0));
                                    } else {
                                        text = new StringBuilder(run.getText(0));
                                    }

                                    if (text == null) {
                                        continue;
                                    }

                                    //check neu param la danh sach
                                    if (key.toLowerCase().contains("list") && text.toString().startsWith("$")
                                            && ( i < runs.size() - 1 &&  runs.get(i + 1).getText(0) != null && text.toString().concat(runs.get(i + 1).getText(0)).toLowerCase().contains("list"))
                                    ) {
                                        int x = i;
                                        while (x < runs.size() - 1) {
                                            x++;
                                            nextRun = runs.get(x);
                                            text.append(nextRun.getText(0));
                                            if(checkSymmetric(text.toString()) && text.toString().contains("{")){
                                                break;
                                            }
                                        }

                                        //check neu key == key in text, else -> continue
                                        if(StringUtils.startsWith(text.toString(), key)){
                                            if(checkSymmetric(text.toString()) && !appendData){
                                                if(listData != null && !listData.isEmpty()){
                                                    replaceListParamDynamic(document, paragraph, listData);
                                                    p += listData.size();
                                                    //remove các run cũ
                                                    while (!paragraph.getRuns().isEmpty()){
                                                        paragraph.removeRun(0);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }


    //gen lại file sau khi sửa ở bước 2// khong dung nua
    public void genFile(Map<String, Map<String, Map<String, String>>> stringMapMaster, XWPFDocument document){
        stringMapMaster.forEach((key, mapParagraph) -> {
            XWPFRun run, nextRun;
            StringBuilder text = null;
            for(int el = 0; el < document.getBodyElements().size(); el++){
                IBodyElement iBodyElement = document.getBodyElements().get(el);
                if(iBodyElement instanceof XWPFParagraph){
                    XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs != null) {
                        for (int i = 0; i < runs.size(); i++) {
                            run = runs.get(i);
                            if(run == null || StringUtils.isEmpty(run.getText(0))) continue;
                            if(text != null && !checkSymmetric(text.toString())){
                                text.append(run.getText(0));
                            } else {
                                text = new StringBuilder(run.getText(0));
                            }

                            if(!runs.isEmpty()) {
                                if(text.toString().contains("$")){
                                    while (!checkVariable(text.toString()) || !checkSymmetric(text.toString())){
                                        if(i == runs.size() - 1){
                                            break;
                                        }
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        i++;
                                    }
                                }
                            }
                            if(key.equals(getKey(text.toString()))){
                                int posOfParagraph = document.getPosOfParagraph(paragraph);
        //                        insert paragraph
                                if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                                    List<Map.Entry<String, Map<String, String>>> entryList = new ArrayList<>(mapParagraph.entrySet());
                                    for (int para = entryList.size() - 1; para >= 0; para--) {
                                        Map.Entry<String, Map<String, String>> entryPara = entryList.get(para);
                                        Map<String, String> stringStringMap = entryPara.getValue();

                                        XWPFParagraph newParagraph = addParagraph(document, paragraph);
                                        // them style cho paragraph
                                        String styleParagraph = stringStringMap.get("styleParagraph");
                                        if(styleParagraph != null){
                                            String align = StringUtils.substringBetween(styleParagraph, "text-align:", ",");
                                            String valfirstLine = StringUtils.substringBetween(styleParagraph, "valfirstLine:", ",");
                                            if(StringUtils.isNotEmpty(valfirstLine)){
                                                //convert px to twips
                                                double valfirstLineTwips = Double.parseDouble(valfirstLine) / convertTwipsToPx;
                                                newParagraph.setFirstLineIndent((int) valfirstLineTwips);
                                            }
                                            if(align != null){
                                                newParagraph.setAlignment(ParagraphAlignment.valueOf(align.toUpperCase()));
                                            }

                                        }

                                        // them style cho run
                                        stringStringMap.entrySet().stream().skip(1).forEach(entry -> {
                                            //neu thay có \n thi xuong dong
                                            String [] textSplitRun = entry.getKey().split("\n");
                                            String style = entry.getValue();
                                            for(int t = 0; t < textSplitRun.length; t++){
                                                String runText = textSplitRun[t];
                                                XWPFRun newRun = newParagraph.createRun();
                                                newRun.setText(runText);

                                                //lay ra style
                                                String boldValueStr = StringUtils.substringBetween(style, "bold:", ",");
                                                boolean boldValue = Boolean.parseBoolean(boldValueStr);

                                                String italicValueStr = StringUtils.substringBetween(style, "italic:", ",");
                                                boolean italicValue = Boolean.parseBoolean(italicValueStr);

                                                String colorValueStr = StringUtils.substringBetween(style, "color:", ",");
                                                String fontSizeValueStr = StringUtils.substringBetween(style, "fontSize:", ",");
                                                String underlineValueStr = StringUtils.substringBetween(style, "underline:", ",");

                                                //convert fontsize px to pt
                                                double fontSizePx = Double.parseDouble(fontSizeValueStr) / convertPtToPx;
                                                newRun.setFontSize((int) fontSizePx);

                                                newRun.setBold(boldValue);
                                                if(!colorValueStr.equals("null")){
                                                    newRun.setColor(colorValueStr);
                                                }
                                                newRun.setItalic(italicValue);
                                                if(StringUtils.isNotEmpty(underlineValueStr)){
                                                    newRun.setUnderline(UnderlinePatterns.valueOf(underlineValueStr));
                                                }
                                                if(textSplitRun.length > 1 && t < textSplitRun.length - 1){
                                                    newRun.addCarriageReturn();
                                                }
                                            }
                                        });
                                    }

                                    document.removeBodyElement(posOfParagraph);
                                } else {
                                    document.removeBodyElement(posOfParagraph);
                                    el--;
                                }
                            }
                        }
                    }
                } else if (iBodyElement instanceof XWPFTable) {
                    XWPFTable tbl = (XWPFTable) iBodyElement;
                    for (XWPFTableRow row : tbl.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (int p = 0; p < cell.getParagraphs().size(); p++) {
                                XWPFParagraph paragraph = cell.getParagraphs().get(p);
                                List<XWPFRun> runs = paragraph.getRuns();
                                if (runs != null) {
                                    for (int i = 0; i < runs.size(); i++) {
                                        run = runs.get(i);
                                        if(StringUtils.isEmpty(run.getText(0))) continue;
                                        if(text != null && !checkSymmetric(text.toString())){
                                            text.append(run.getText(0));
                                        } else {
                                            text = new StringBuilder(run.getText(0));
                                        }

                                        if(!runs.isEmpty()) {
                                            if(text.toString().contains("$")){
                                                while (!checkVariable(text.toString()) || !checkSymmetric(text.toString())){
                                                    if(i == runs.size() - 1){
                                                        break;
                                                    }
                                                    nextRun = runs.get(i + 1);
                                                    text.append(nextRun.getText(0));
                                                    i++;
                                                }
                                            }
                                        }
                                        if(key.equals(getKey(text.toString())) && checkSymmetric(text.toString())){
                                            cell.removeParagraph(p);
                                            //insert paragraph
                                            if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                                                XWPFParagraph newParagraph = cell.addParagraph();
                                                mapParagraph.forEach((s, stringStringMap) -> {
                                                    // them style cho paragraph
                                                    String styleParagraph = stringStringMap.get("styleParagraph");
                                                    if(styleParagraph != null){
                                                        String align = StringUtils.substringBetween(styleParagraph, "text-align:", ",");
                                                        String valfirstLine = StringUtils.substringBetween(styleParagraph, "valfirstLine:", ",");
                                                        if(StringUtils.isNotEmpty(valfirstLine)){
                                                            //convert px to twips
                                                            double valfirstLineTwips = Double.parseDouble(valfirstLine) / convertTwipsToPx;
                                                            newParagraph.setFirstLineIndent((int) valfirstLineTwips);
                                                        }
                                                        if(align != null){
                                                            newParagraph.setAlignment(ParagraphAlignment.valueOf(align.toUpperCase()));
                                                        }
                                                    }

                                                    //them style cho tung run
                                                    stringStringMap.entrySet().stream().skip(1).forEach(entry -> {
                                                        //neu thay có \n thi xuong dong
                                                        String [] textSplitRun = entry.getKey().split("\n");
                                                        String style = entry.getValue();
                                                        for(String runText : textSplitRun){
                                                            XWPFRun newRun = newParagraph.createRun();
                                                            if(textSplitRun.length > 1){
                                                                newRun.addBreak();
                                                            }
                                                            newRun.setText(runText);

                                                            //lay ra style
                                                            String boldValueStr = StringUtils.substringBetween(style, "bold:", ",");
                                                            boolean boldValue = Boolean.parseBoolean(boldValueStr);

                                                            String italicValueStr = StringUtils.substringBetween(style, "italic:", ",");
                                                            boolean italicValue = Boolean.parseBoolean(italicValueStr);

                                                            String colorValueStr = StringUtils.substringBetween(style, "color:", ",");
                                                            String fontSizeValueStr = StringUtils.substringBetween(style, "fontSize:", ",");
                                                            String underlineValueStr = StringUtils.substringBetween(style, "underline:", ",");

                                                            //convert fontsize px to pt
                                                            double fontSizePx = Double.parseDouble(fontSizeValueStr) / convertPtToPx;
                                                            newRun.setFontSize((int) fontSizePx);

                                                            newRun.setBold(boldValue);
                                                            if(!colorValueStr.equals("null")){
                                                                newRun.setColor(colorValueStr);
                                                            }
                                                            newRun.setItalic(italicValue);
                                                            if(StringUtils.isNotEmpty(underlineValueStr)){
                                                                newRun.setUnderline(UnderlinePatterns.valueOf(underlineValueStr));
                                                            }
                                                        }
                                                    });
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        });
    }

    //gen lại file sau khi sửa ở bước 2 v2
    public void genFileV2(Map<String, List<GenFileDTO>> stringListMap, XWPFDocument document){
        stringListMap.forEach((key, genFileDTOS) -> {
            XWPFRun run, nextRun;
            StringBuilder text = null;
            for(int el = 0; el < document.getBodyElements().size(); el++){
                IBodyElement iBodyElement = document.getBodyElements().get(el);
                if(iBodyElement instanceof XWPFParagraph){
                    XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs != null) {
                        for (int i = 0; i < runs.size(); i++) {
                            run = runs.get(i);
                            if(run == null || StringUtils.isEmpty(run.getText(0))) continue;
                            if(text != null && !checkSymmetric(text.toString())){
                                text.append(run.getText(0));
                            } else {
                                text = new StringBuilder(run.getText(0));
                            }

                            if(!runs.isEmpty()) {
                                if(text.toString().contains("$")){
                                    while (!checkVariable(text.toString()) || !checkSymmetric(text.toString())){
                                        if(i == runs.size() - 1){
                                            break;
                                        }
                                        nextRun = runs.get(i + 1);
                                        text.append(nextRun.getText(0));
                                        i++;
                                    }
                                }
                            }
                            if(key.equals(getKey(text.toString()))){
                                int posOfParagraph = document.getPosOfParagraph(paragraph);
                                //insert paragraph
                                if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                                    for (int para = genFileDTOS.size() - 1; para >= 0; para--) {
                                        GenFileDTO genFileDTO = genFileDTOS.get(para);
                                        XWPFParagraph newParagraph = addParagraph(document, paragraph);
                                        // them style cho paragraph
                                        String styleParagraph = genFileDTO.getStyleParagraph();
                                        if(styleParagraph != null){
                                            String align = StringUtils.substringBetween(styleParagraph, "text-align:", ",");
                                            String valfirstLineStr = StringUtils.substringBetween(styleParagraph, "valfirstLine:", ",");
                                            if(StringUtils.isNotEmpty(valfirstLineStr)){
                                                //convert px to twips
                                                String valfirstLine = valfirstLineStr.substring(0, valfirstLineStr.length() - 2);
                                                double valfirstLineTwips = Double.parseDouble(valfirstLine) / convertTwipsToPx;
                                                newParagraph.setFirstLineIndent((int) valfirstLineTwips);
                                            }
                                            if(align != null){
                                                newParagraph.setAlignment(ParagraphAlignment.valueOf(align.toUpperCase()));
                                            }
                                        }

                                        // them style cho run
                                        genFileDTO.getChildrens().stream().forEach(runDTO -> {
                                            //neu thay có \n thi xuong dong
                                            String [] textSplitRun = runDTO.getTitle().split("\n");
                                            String style = runDTO.getStyleParagraph();
                                            for(int t = 0; t < textSplitRun.length; t++){
                                                String runText = textSplitRun[t];
                                                XWPFRun newRun = newParagraph.createRun();
                                                newRun.setText(runText);

                                                //lay ra style
                                                String boldValueStr = StringUtils.substringBetween(style, "bold:", ",");
                                                if(StringUtils.isNotEmpty(boldValueStr)){
                                                    boolean boldValue = Boolean.parseBoolean(boldValueStr);
                                                    newRun.setBold(boldValue);
                                                }

                                                String italicValueStr = StringUtils.substringBetween(style, "italic:", ",");
                                                if(StringUtils.isNotEmpty(italicValueStr)){
                                                    boolean italicValue = Boolean.parseBoolean(italicValueStr);
                                                    newRun.setItalic(italicValue);
                                                }

                                                String fontSizeValueStr = StringUtils.substringBetween(style, "fontSize:", ",");
                                                if(StringUtils.isNotEmpty(fontSizeValueStr)){
                                                    //convert fontsize px to pt
                                                    String fontSizeValue = fontSizeValueStr.substring(0, fontSizeValueStr.length() - 2);
                                                    if(StringUtils.isNotEmpty(fontSizeValue)){
                                                        double fontSizePx = Double.parseDouble(fontSizeValue) / convertPtToPx;
                                                        newRun.setFontSize((int) fontSizePx);
                                                    }
                                                }

                                                String strikeStr = StringUtils.substringBetween(style, "strike:", ",");
                                                if(StringUtils.isNotEmpty(strikeStr)){
                                                    boolean strike = Boolean.parseBoolean(strikeStr);
                                                    newRun.setStrikeThrough(strike);
                                                }

                                                String subscript = StringUtils.substringBetween(style, "subscript:", ",");
                                                if(StringUtils.isNotEmpty(subscript)){
                                                    newRun.setSubscript(VerticalAlign.valueOf(subscript.toUpperCase()));
                                                }

                                                String underlineValueStr = StringUtils.substringBetween(style, "underline:", ",");
                                                if(StringUtils.isNotEmpty(underlineValueStr)){
                                                    newRun.setUnderline(UnderlinePatterns.valueOf(underlineValueStr.toUpperCase()));
                                                }

                                                if(textSplitRun.length > 1 && t < textSplitRun.length - 1){
                                                    newRun.addCarriageReturn();
                                                }
                                            }
                                        });
                                    }
                                    document.removeBodyElement(posOfParagraph);
                                } else {
                                    document.removeBodyElement(posOfParagraph);
                                    el--;
                                }
                            }
                        }
                    }
                } else if (iBodyElement instanceof XWPFTable) {
                    XWPFTable tbl = (XWPFTable) iBodyElement;
                    for (XWPFTableRow row : tbl.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (int p = 0; p < cell.getParagraphs().size(); p++) {
                                XWPFParagraph paragraph = cell.getParagraphs().get(p);
                                List<XWPFRun> runs = paragraph.getRuns();
                                if (runs != null) {
                                    for (int i = 0; i < runs.size(); i++) {
                                        run = runs.get(i);
                                        if(StringUtils.isEmpty(run.getText(0))) continue;
                                        if(text != null && !checkSymmetric(text.toString())){
                                            text.append(run.getText(0));
                                        } else {
                                            text = new StringBuilder(run.getText(0));
                                        }

                                        if(!runs.isEmpty()) {
                                            if(text.toString().contains("$")){
                                                while (!checkVariable(text.toString()) || !checkSymmetric(text.toString())){
                                                    if(i == runs.size() - 1){
                                                        break;
                                                    }
                                                    nextRun = runs.get(i + 1);
                                                    text.append(nextRun.getText(0));
                                                    i++;
                                                }
                                            }
                                        }
                                        if(key.equals(getKey(text.toString())) && checkSymmetric(text.toString())){
                                            cell.removeParagraph(p);
                                            //insert paragraph
                                            if(checkVariable(text.toString()) && checkSymmetric(text.toString())){
                                                XWPFParagraph newParagraph = cell.addParagraph();
                                                genFileDTOS.forEach(genFileDTO -> {
                                                    // them style cho paragraph
                                                    String styleParagraph = genFileDTO.getStyleParagraph();
                                                    if(styleParagraph != null){
                                                        String align = StringUtils.substringBetween(styleParagraph, "text-align:", ",");
                                                        String valfirstLineStr = StringUtils.substringBetween(styleParagraph, "valfirstLine:", ",");
                                                        if(StringUtils.isNotEmpty(valfirstLineStr)){
                                                            //convert px to twips
                                                            String valfirstLine = valfirstLineStr.substring(0, valfirstLineStr.length() - 2);
                                                            double valfirstLineTwips = Double.parseDouble(valfirstLine) / convertTwipsToPx;
                                                            newParagraph.setFirstLineIndent((int) valfirstLineTwips);
                                                        }
                                                        if(align != null){
                                                            newParagraph.setAlignment(ParagraphAlignment.valueOf(align.toUpperCase()));
                                                        }
                                                    }

                                                    // them style cho run
                                                    genFileDTO.getChildrens().stream().forEach(runDTO -> {
                                                        //neu thay có \n thi xuong dong
                                                        String [] textSplitRun = runDTO.getTitle().split("\n");
                                                        String style = runDTO.getStyleParagraph();
                                                        for(int t = 0; t < textSplitRun.length; t++){
                                                            String runText = textSplitRun[t];
                                                            XWPFRun newRun = newParagraph.createRun();
                                                            newRun.setText(runText);

                                                            //lay ra style
                                                            String boldValueStr = StringUtils.substringBetween(style, "bold:", ",");
                                                            if(StringUtils.isNotEmpty(boldValueStr)){
                                                                boolean boldValue = Boolean.parseBoolean(boldValueStr);
                                                                newRun.setBold(boldValue);
                                                            }

                                                            String italicValueStr = StringUtils.substringBetween(style, "italic:", ",");
                                                            if(StringUtils.isNotEmpty(italicValueStr)){
                                                                boolean italicValue = Boolean.parseBoolean(italicValueStr);
                                                                newRun.setItalic(italicValue);
                                                            }

                                                            String fontSizeValueStr = StringUtils.substringBetween(style, "fontSize:", ",");
                                                            if(StringUtils.isNotEmpty(fontSizeValueStr)){
                                                                //convert fontsize px to pt
                                                                String fontSizeValue = fontSizeValueStr.substring(0, fontSizeValueStr.length() - 2);
                                                                if(StringUtils.isNotEmpty(fontSizeValue)){
                                                                    double fontSizePx = Double.parseDouble(fontSizeValue) / convertPtToPx;
                                                                    newRun.setFontSize((int) fontSizePx);
                                                                }
                                                            }

                                                            String strikeStr = StringUtils.substringBetween(style, "strike:", ",");
                                                            if(StringUtils.isNotEmpty(strikeStr)){
                                                                boolean strike = Boolean.parseBoolean(strikeStr);
                                                                newRun.setStrikeThrough(strike);
                                                            }

                                                            String subscript = StringUtils.substringBetween(style, "subscript:", ",");
                                                            if(StringUtils.isNotEmpty(subscript)){
                                                                newRun.setSubscript(VerticalAlign.valueOf(subscript.toUpperCase()));
                                                            }

                                                            String underlineValueStr = StringUtils.substringBetween(style, "underline:", ",");
                                                            if(StringUtils.isNotEmpty(underlineValueStr)){
                                                                newRun.setUnderline(UnderlinePatterns.valueOf(underlineValueStr.toUpperCase()));
                                                            }

                                                            if(textSplitRun.length > 1 && t < textSplitRun.length - 1){
                                                                newRun.addCarriageReturn();
                                                            }
                                                        }
                                                    });
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        });
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

    //get value từ 1 đoạn text có cả key và value
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
