<%
    Class count_execution_time
        Dim start_time 

        'Initialization and destruction'
        Sub class_initialize()
        End Sub 

        Sub class_terminate()
            start_time = Nothing
        End Sub 

        'Set starting count time 
        Public Function set_start_time()
            start_time = Timer()
        End Function 

        'Retrieve the execution time in seconds 
        Public Function get_execution_time()
            Dim temp_time
            temp_time = Timer()
            get_execution_time = (temp_time - start_time)
        End Function 
    End Class 
%>