# Creon 주가 수집 & 백테스트 프로그램

![GIF 2020-09-16 수 오후 8-43-40](https://user-images.githubusercontent.com/15887982/93333254-030cb480-f85e-11ea-9884-fbbbb5977850.gif)

## 사용언어

python

<br>

## 프로젝트 설명

크레온 라이브러리를 활용한 주가 데이터 수집 및 백테스트 파일입니다.<br>

DB 업데이트(주가 데이터가 없는 경우 새로 받기)하여 주가를 저장합니다.<br>

<br>

## 주요기능 (프로젝트의 모든 기능을 혼자 구현했습니다)

- 1분봉, 5분봉 등 각종 분봉 및 시봉, 일봉 수집
- 투자 전략(돌파 전략) 백테스트 기능

<br>

## 기술적 문제 해결

- 주가 데이터를 어디서 받을 것인가
  - Creon Library 이용
  - Creon에서만 1분봉 2년치 데이터를 제공하기 때문에 선택함
- 수정주가 처리 문제
  - 일봉 데이터는 1999년부터 받는 것이 가능함
  - 따라서 일봉 데이터를 통해 수정주가 존재 여부 확인함
  - 이후 다른 봉들은 일봉을 기준으로 수정주가 처리
