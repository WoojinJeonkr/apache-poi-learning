package com.example.apachepoilearning.domain.upload.service;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

public class SheetHandler extends DefaultHandler {

    // 공유 문자열 테이블 (SharedStringsTable) - 엑셀 파일 내의 중복되는 문자열들을 효율적으로 저장하는 테이블
    // 엑셀 파일에서 "s" 타입의 셀(문자열 셀)은 이 테이블의 인덱스를 참조하여 실제 문자열 값을 가져옵니다.
    private final SharedStringsTable sst;
    // 현재 읽고 있는 요소의 내용을 임시로 저장하는 변수
    private String lastContents = "";
    // 현재 셀이 문자열 타입인지 여부를 나타내는 플래그
    private boolean isString = false;

    /**
     * SheetHandler의 생성자
     * @param sst 공유 문자열 테이블 객체
     */
    public SheetHandler(SharedStringsTable sst) {
        this.sst = sst;
    }

    /**
     * XML 요소의 시작 태그를 만났을 때 호출
     * 예를 들어, <c> (셀 시작) 또는 <row> (행 시작) 태그 등 처리
     * @param uri 네임스페이스 URI
     * @param localName 접두사 없는 로컬 이름
     * @param name 접두사가 붙은 정규화된 이름 (qName)
     * @param attributes 요소의 속성들
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        // 현재 요소가 "c" (셀) 태그인 경우
        if ("c".equals(name)) {
            // 셀의 't' (type) 속성 값을 가져옴
            String type = attributes.getValue("t");
            // 셀 타입이 "s" (shared string) 인지 확인하여 isString 플래그 설정
            // "s" 타입은 SharedStringsTable에서 실제 문자열 찾아야함
            isString = "s".equals(type);
        }
        // 새로운 요소가 시작될 때마다 lastContents 초기화
        lastContents = "";
    }

    /**
     * XML 요소 내의 문자 데이터를 만났을 때 호출
     * 예를 들어, <v>123</v> 에서 "123" 부분을 읽을 때 사용
     * @param ch 문자 배열
     * @param start 문자 데이터의 시작 인덱스
     * @param length 문자 데이터의 길이
     */
    @Override
    public void characters(char[] ch, int start, int length) {
        // 읽은 문자 데이터를 lastContents 추가
        // 한 번에 모든 문자열이 들어오지 않을 수 있으므로 += 사용
        lastContents += new String(ch, start, length);
    }

    /**
     * XML 요소의 종료 태그를 만났을 때 호출
     * 예를 들어, </c> (셀 종료) 또는 </row> (행 종료) 태그 등 처리
     * @param uri 네임스페이스 URI
     * @param localName 접두사 없는 로컬 이름
     * @param name 접두사가 붙은 정규화된 이름 (qName)
     */
    @Override
    public void endElement(String uri, String localName, String name) {
        // 현재 요소가 "v" (셀 값) 태그인 경우
        if ("v".equals(name)) {
            // 셀이 문자열 타입(isString이 true)인 경우
            if (isString) {
                // lastContents는 SharedStringsTable의 인덱스이므로 정수로 변환
                int idx = Integer.parseInt(lastContents);
                // 해당 인덱스를 사용하여 SharedStringsTable에서 실제 문자열 값 가져옴
                lastContents = sst.getItemAt(idx).getString();
            }
            // 현재 셀의 값을 출력하고 탭으로 구분
            System.out.print(lastContents + "\t");
        }

        // 현재 요소가 "row" (행) 태그인 경우
        if ("row".equals(name)) {
            // 한 행의 처리가 끝났으므로 줄바꿈
            System.out.println();
        }
    }

}
