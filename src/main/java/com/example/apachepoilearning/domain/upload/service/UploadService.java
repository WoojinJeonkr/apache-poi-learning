package com.example.apachepoilearning.domain.upload.service;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * 엑셀 파일 업로드 및 처리를 담당하는 서비스 클래스
 * Apache POI 라이브러리를 사용하여 엑셀 파일을 읽는 두 가지 방식을 보여줍니다
 * 1. XSSFWorkbook 방식 (객체 모델 기반): 작거나 중간 크기의 파일에 적합하며 사용하기 쉽습니다.
 * 2. SAX 방식 (이벤트 기반): 대용량 파일 처리에 적합하며 메모리 효율적입니다.
 */
@Service
public class UploadService {

    /**
     * 엑셀 파일(XLSX)을 XSSFWorkbook (객체 모델) 방식으로 업로드하고 처리합니다.
     * 이 방식은 파일을 모두 메모리에 로드하므로, 파일 크기가 작거나 중간일 때 적합합니다.
     *
     * @param file 업로드된 MultipartFile 객체 (스프링 웹 환경에서 파일 업로드를 통해 전달받음)
     * @throws IOException 파일 입출력 중 발생할 수 있는 예외
     */
    public void uploadXlsx(MultipartFile file) throws IOException {

        // 업로드된 파일의 원본 이름 출력
        System.out.println("업로드 파일명: " + file.getOriginalFilename());

        // 업로드한 파일을 읽기 위한 InputStream 얻기
        InputStream inputStream = file.getInputStream();

        // InputStream 으로부터 XSSFWorkbook 객체 (엑셀 워크북) 생성
        // 이 시점에서 전체 엑셀 파일이 메모리에 로드됩니다.
        Workbook workbook = new XSSFWorkbook(inputStream);

        // --- 시트 정보 확인 ---
        // 워크북의 전체 시트 개수 출력
        int sheetCount = workbook.getNumberOfSheets(); // 시트 개수
        System.out.println("총 시트 개수: " + sheetCount);

        // 각 시트의 이름 출력
        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) { // 시트 이름 출력
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println("시트 이름 (" + i + "번째): " + sheet.getSheetName());
        }

        // --- 특정 시트 읽기 및 데이터 처리 ---
        // 첫 번째 시트를 선택합니다. (일반적으로 모든 시트를 반복문으로 처리)
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println("\n첫 번째 시트 내용 출력:");

        // 선택된 시트의 모든 행을 반복하며 처리
        for (Row row : sheet) {
            // 각 행의 모든 셀을 반복하며 처리
            for (Cell cell : row) {
                // 셀의 내용을 콘솔에 출력 (기본 toString() 호출)
                System.out.println(cell + "\t"); // 탭으로 구분하여 출력
            }
            // 한 행의 처리가 끝나면 줄바꿈
            System.out.println();
        }
    }

    /**
     * 엑셀 파일(XLSX)을 SAX 파싱 방식으로 업로드하고 처리합니다.
     * 이 방식은 파일을 이벤트 기반으로 처리하여 대용량 파일에 대한 메모리 효율성이 뛰어납니다.
     *
     * @param file 업로드된 MultipartFile 객체
     * @throws Exception 파싱 및 파일 처리 중 발생할 수 있는 예외
     */
    public void uploadSAXXlsx(MultipartFile file) throws Exception {
        System.out.println("\nSAX 방식 엑셀 파일 처리 시작...");

        // OPCPackage 열기: XLSX 파일은 ZIP 압축 파일 형식이며, OPC 패키지로 관리됩니다.
        // 이를 통해 파일 내부의 XML 스트림에 직접 접근할 수 있습니다.
        OPCPackage pkg = OPCPackage.open(file.getInputStream());
        // XSSFReader 생성: OPC 패키지에서 엑셀 관련 XML 스트림(시트, 공유 문자열 등)을 읽기 위한 리더
        XSSFReader xssfReader = new XSSFReader(pkg);

        // --- 시트 정보 출력 ---
        // getSheetNames 유틸리티 메서드를 사용하여 시트 이름 목록 가져오기
        List<String> sheetNames = getSheetNames(pkg);
        System.out.println("SAX 방식 - 총 시트 개수: " + sheetNames.size());
        System.out.println("SAX 방식 - 시트 이름");
        for (String name : sheetNames) {
            System.out.println(name);
        }

        // --- 시트 데이터 스트림 및 공유 문자열 테이블 가져오기 ---
        // getSheetsData()는 시트 데이터 스트림들의 Iterator를 반환합니다.
        // next()를 호출하여 첫 번째 시트의 InputStream을 가져옵니다.
        InputStream sheetStream = xssfReader.getSheetsData().next();

        // SharedStringsTable 가져오기: 엑셀 파일 내의 모든 공유 문자열(텍스트)을 저장하는 테이블입니다.
        // SAX 파싱 시 셀 값이 문자열 인덱스로 되어있으므로, 이 테이블에서 실제 문자열을 조회해야 합니다.
        SharedStrings sst = xssfReader.getSharedStringsTable();
        SharedStringsTable sstTable = (SharedStringsTable) sst;

        // --- XML 파서 및 핸들러 설정 ---
        // SAXParserFactory를 통해 XMLReader 인스턴스 생성
        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        // 커스텀 핸들러 (SheetHandler)를 파서에 설정합니다.
        // SheetHandler는 XML을 파싱하면서 발생하는 이벤트(요소 시작/종료, 문자 데이터 등)를 처리합니다.
        parser.setContentHandler(new SheetHandler(sstTable));

        // --- 파싱 시작 ---
        // InputSource를 사용하여 시트 데이터 스트림을 파서에 전달하여 파싱을 시작합니다.
        // 파싱이 진행되면서 SheetHandler의 메서드들이 호출되어 셀 데이터를 처리합니다.
        parser.parse(new InputSource(sheetStream));
    }

    /**
     * OPCPackage 에서 엑셀 워크북의 XML 스트림을 파싱하여 시트 이름 목록을 추출합니다.
     * 이 메서드도 SAX 파싱 방식을 사용하여 메모리 효율적으로 시트 이름을 가져옵니다.
     *
     * @param pkg 열려있는 OPCPackage 객체
     * @return 시트 이름들을 담은 List<String>
     * @throws Exception XML 파싱 및 파일 접근 중 발생할 수 있는 예외
     */
    private List<String> getSheetNames(OPCPackage pkg) throws Exception {

        List<String> names = new ArrayList<>();
        // "/xl/workbook.xml" 경로는 엑셀 파일(ZIP) 내부에 워크북 구조 정보를 담고 있는 XML 파일입니다.
        // Pattern.compile을 사용하여 해당 경로의 스트림을 가져옵니다.
        InputStream workbookXml = pkg.getPartsByName(Pattern.compile("/xl/workbook.xml")).get(0).getInputStream();

        // XMLReader 생성 및 핸들러 설정 (익명 클래스 사용)
        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        parser.setContentHandler(new DefaultHandler() { // 익명 내부 클래스로 DefaultHandler 구현

            /**
             * XML 요소의 시작 태그를 만났을 때 호출됩니다.
             * <sheet> 태그를 찾아 'name' 속성 값을 시트 이름으로 추출합니다.
             */
            @Override
            public void startElement(String uri, String localName, String name, Attributes attributes) {
                // 현재 요소 이름이 "sheet"인 경우
                if ("sheet".equals(name)) {
                    // "name" 속성 값을 가져와 리스트에 추가
                    names.add(attributes.getValue("name"));
                }
            }
        });

        // workbook.xml 스트림을 파싱하여 시트 이름을 추출합니다.
        parser.parse(new InputSource(workbookXml));
        return names;
    }
}
