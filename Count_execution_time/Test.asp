<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="count_execution_time.class.asp"-->
<%
    Dim time
    Dim i,a
    Set time = new count_execution_time

    time.set_start_time()

    'Simulate some work
    For i = 1 To 100000
    a = a + i
    Next

    Response.write(time.get_execution_time())
%>
