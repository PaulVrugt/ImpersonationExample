<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="ImpersonationTest._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="jquery.min.js"></script>
    <script type="text/javascript">

        function go()
        {
            for (var intI = 0; intI < 100; intI++)
            {
                $.ajax({
                    method: "POST",
                    url: "Page.aspx?count=" + intI,
                    data: '{"test":"aap","harrie":"nak","gerard":"harrie"}',
                    success: function (result)
                    {
                        document.getElementById('divResult').innerHTML += "<br/> " + result;
                    },
                    error: function ()
                    {
                        debugger;
                    }
                });
            }
        }

        $(document).ready(go);

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div id="divResult">

        </div>
    </form>
</body>
</html>
