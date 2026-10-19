# 지역별 가정용 소주 구매액 1위는?

**2026년 5월 MAT 기준, 공개된 5개 권역의 가정용 소주류 구매액 TOP 7을 비교한 데이터 리터러시 프로젝트**입니다.

## 핵심 결과

- 수도권·충청권·전라권·강원권 1위: 참이슬 후레쉬
- 경상권 1위: 진로
- 같은 1위가 나타난 권역: 5개 중 4개

이 순위는 가정용 소주류의 **구매액 순위**입니다. 소비량, 판매량, 시장점유율, 외식 구매, 전체 소비자 선호도와 같지 않습니다. 구매액 원수치와 상세 패널 정보는 원문에 공개되지 않았습니다.

## 결과물

- `presentation/soju_regional_rank_2026.pptx`: 10장, PowerPoint에서 편집 가능한 발표자료
- `presentation/speaker_notes.md`: 발표자 노트
- `docs/index.html`: 방향키·버튼·진행률·전체 화면·터치 조작 및 차트 모션이 있는 웹 슬라이드
- `docs/data/regional_top7_2026_mat.csv`: 웹 자료 페이지에서 사용하는 전사 데이터
- `data/regional_top7_2026_mat.csv`: 6개 그룹(전국+5권역) × TOP 7, 총 42행
- `data/regional_representative_vs_rank1_2026.csv`: 권역별 1위 요약
- `data/sources_2026.csv`: 출처·기준 기간·자료 한계
- `analysis/insights.md`, `analysis/data_quality_report.md`: 해석 및 품질 점검
- `scripts/build_deck.mjs`: 데이터로부터 PPTX와 HTML을 재생성하는 코드

## 재생성 방법

Node.js 20 이상 필요:

```powershell
npm install
node scripts/build_deck.mjs
```

웹 발표 자료는 `docs/index.html`을 브라우저에서 열면 됩니다. anime.js는 CDN을 사용합니다. 오프라인에서는 애니메이션이 작동하지 않을 수 있지만 슬라이드 내용과 조작은 유지됩니다.

## 수집 및 분석 방법

트렌드모니터의 공개 페이지와 첨부 인포그래픽을 직접 확인해 순위만 전사했습니다. 기준 기간은 페이지 표기대로 2026년 5월 MAT(최근 12개월)입니다. 원문은 구매 추정액이 실제 결제 기반 추정치이며 공식 매출액 또는 소비자 선호도 순위와 다를 수 있다고 고지합니다.

순위 자료이므로 순위형 표와 지역 카드로 비교했습니다. 구매액 격차가 공개되지 않아 이를 막대 길이나 시장 점유율처럼 시각화하지 않았습니다. 관련 지역 브랜드 상품도 해당 권역 TOP 7 안에서의 순위로만 기록했습니다.

## 주요 출처

- 트렌드모니터·엠브레인 구매딥데이터 DD.B, [지역별 최애 소주 순위 TOP7](https://www.trendmonitor.co.kr/tmweb/trend/etcTrend/detail.do?bIdx=36&code=DDB&trendType=ranking), 게시 2026-07-02, 기준 2026년 5월 MAT.
- [Puppia 보조 해설](https://hobby.puppia.com/posts/or-mv9fdjdl)은 맥락 참고용이며 순위 전사 및 결론의 원자료로 사용하지 않았습니다.

## 한계

5개 권역만 제공되어 제주도는 별도 분석할 수 없습니다. 구매액 원수치·판매량·점유율·패널 규모·표집과 가중 세부 방법이 없고, 가정용 구매만 다룹니다. 한 기간의 순위로 지역 정체성이나 브랜드 선호 원인을 설명할 수 없습니다. 2024~25년 언론 보도 기반 회사 점유율과 과거 조사 자료는 단위와 기간이 달라 이번 순위와 합산하지 않았습니다.

## 발표 URL

배포 후 GitHub Pages 주소를 여기에 기록합니다.
