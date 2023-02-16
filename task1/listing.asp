<!DOCTYPE html>
<%@ language = vbscript CODEPAGE=65001%>

<html>
    <%pagename ="all eligible NH employees"%>
    <!--#include file='mainheader.asp'-->
    <!--#include file='functions.asp'-->

    <%
    If Request.ServerVariables("REQUEST_METHOD") = "POST" THEN

        If Request.querystring("action") = "search" THEN
            ELGIBLE = request("eligible")
        End If
    End If

    
    %>
    <body>
    <div class="page-container row-fluid">
        <div class="page-content">
            <div class="content">
                <div class="page-title">
                    <div class="row">
                        <div class="col-md-12">
                            <h3>Listing Page</h3>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class="row">
        <ul class="nav nav-tabs" role="tablist">
            <li class="active">
                <a href="#tab1helloWorld" role="tab" data-toggle="tab">
                    <span class="semi-bold">All Eligible</span>
                </a>
            </li>
        </ul>
        <div class="tab-content">
            <div class="col-md-4 p-t-10 p-b-10 p-l-35">
                <form class="" action="listing.asp?action=search" method="POST">
                    <div class="row">
                        <div class="col-md-4">
                            <select class="form-control" id="eligible" name="eligible">
                                <option value="0" <%if ELGIBLE=0 then%> selected <%end if%>> Non-Eliglible</option>
                                <option value="1" <%if ELGIBLE=1 then%> selected <%end if%>> Eliglible</option>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <input class="btn btn-primary form-control" value="search" type="submit">
                        </div>
                    </div>
                </form>
            </div>
            <%if ELGIBLE = 0 then%>
            <div class="tab-pane active" id="tab1helloWorld">
                <%
                    strsql = getEligible()
                    set rs = scn.execute(strsql)
                %>
            <div class="col-md-12" style="overflow: auto;" id="tablewrapper">
                <table class="table table-bordered no-more-tables" id="eligibleTable">
                    <thead>
                        <tr>
                            
                            <th style="vertical-align: text-top;">emplid</th>
                            <th style="vertical-align: text-top;"></th>
                            <th style="vertical-align: text-top;"></th>
                            <th style="vertical-align: text-top;"></th>
                            <th style="vertical-align: text-top;"></th>
                            <th style="vertical-align: text-top;"></th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                            if not rs.eof then
                            do until rs.eof
                        %>
                        <tr>
                            <td>
                                <%=rs("emplid")%>
                            </td>
                            <td>
                                <%=rs("fullname")%>
                            </td>
                        </tr><%
                        rs.movenext
                        loop
                        end if
                        %>
                    </tbody>
                </table>

    </div>

            
    </body>
</html>