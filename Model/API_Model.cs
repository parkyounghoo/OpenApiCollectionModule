namespace Open_Api_Collection_Module.Model
{
    internal class API_Model
    {
        public string API_NAME { get; set; } //오픈 API 명
        public string API_Purpose { get; set; } //사용목적
        public string API_URL { get; set; } //호출 URL
        public string API_Response { get; set; } //응답형태
        public string TABLE_NAME { get; set; } //테이블 명
        public string API_Proc_Name { get; set; } //프로시저 명
        public string API_CreateDt { get; set; } //생성날짜
        public string API_CREATE_TABLE { get; set; } //테이블 쿼리
        public string API_CREATE_PROC { get; set; } //프로시저 쿼리
        public string API_Date_YN { get; set; } //날짜 값 YN
        public string API_S_Date_Parameter { get; set; } //시작날짜 Parameter 값
        public string API_E_Date_Parameter { get; set; } //종료날짜 Parameter 값
        public string API_PageResult_YN { get; set; } //페이지 번호 YN
        public string API_PageResult_Name { get; set; } //페이지 번호 Parameter 값
        public string API_PageOfRows_Name { get; set; } //페이지당 건수 Parameter 값
        public string API_TotalCount_Name { get; set; } //전체 결과 수 Parameter  값
        public string API_Data_Collection { get; set; } //데이터 수집처
        public string API_Data_Service { get; set; } //제공기관
        public string API_Collection_Cycle { get; set; } //데이터 수집 주기
        public string API_Collection_Cycle_Date { get; set; } //데이터 수집 날짜
    }

    class TableModel
    {
        public string ColumnName { get; set; }
        public string ColumnType { get; set; }
        public string ColumnSize { get; set; }
        public string ColumnDescription { get; set; }
        public string Value { get; set; }
    }
}