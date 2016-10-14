<%@ page language="java" contentType="text/html; charset=Windows-31J" pageEncoding="Windows-31J"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>ExcelPOISample</title>
</head>
<body>

  <br /><br />
  エクセルファイルを出力します。
  <br /><br />

  <form method="POST" action="/ExcelPOISampleWeb/ExcelOut">
    出力ファイル名：<input type="text" name="fileName" size="10">.xlsx
    <br /><br />
    <input type="submit" value="Excel出力">
    <br />
  </form>

</body>
</html>