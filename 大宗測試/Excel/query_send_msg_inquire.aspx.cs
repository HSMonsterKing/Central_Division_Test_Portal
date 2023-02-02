<% @Page Language="JScript" aspcompat="true" %>
<!--#include file="../../common/sjs/aspxUtil.js"-->
<!--#include file="../../common/sjs/xmlUtil.js"-->
<!--#include file="../../common/sjs/common.js"-->
<!--#include file="../../common/sjs/error_handle.js"-->
<%
var reqDoc=new XMLDoc(true);
var Result=new XMLDoc(true);
var ConnectStr=Application("ConnectStr");
Result.loadXML("<RESPONSE></RESPONSE>");

var Server_Date=new Date();
var Server_Year=Server_Date.getFullYear()-1911;
var Server_Month=Server_Date.getMonth()+1;
var Server_Day=Server_Date.getDate();
var Server_Hour=Server_Date.getHours();
var Server_Min=Server_Date.getMinutes();
var Server_Sec=Server_Date.getSeconds();
Result.setAttribute("Y",Server_Year);
Result.setAttribute("M",Server_Month);
Result.setAttribute("D",Server_Day);
Result.setAttribute("H",Server_Hour);
Result.setAttribute("MIN",Server_Min);
Result.setAttribute("S",Server_Sec);

loadReqDoc(reqDoc);
var Conn=Server.CreateObject("ADODB.Connection");
var MAX_QUERY_NUM=0;	
var doc_no_qry_flag = 0;	
	
try
{
	Conn.Open(ConnectStr);
	var CMD = Server.CreateObject("Adodb.Command");
	CMD.ActiveConnection = Conn;
	
	CMD.CommandType = 1;
	
	
	if (!(reqDoc.toChild("/REQUEST/MAX_QUERY_NUM")))
	{
		Result.appendChild("RETURN",1);		//1:無系統資訊
	}else{
		MAX_QUERY_NUM=reqDoc.getText();
    if (!(reqDoc.toChild("/REQUEST/QRY_BY_DOC_NO")))
		Result.appendChild("RETURN",15);	//15:無doc_no_qry_flag
	else
		doc_no_qry_flag = reqDoc.getText();
	
		CMD.CommandText=Make_SQL_COMMAND();
//Result.appendChild("SQL_COM",CMD.CommandText);
//Result.save(Server.MapPath("phenix.xml") );
		
		var Rec=CMD.Execute();
		var IID=0;
		if (Rec.State==1)
		{
			var apID = 0
			while (!(Rec.EOF))
			{
				
				if(apID==Rec.Fields("AP_ID").Value){
//					Result.toParent();
					Rec.MoveNext();
					continue;
				}

				Result.appendChild("ITEM");
				Result.toLastChild();

				apID = Rec.Fields("AP_ID").Value
		
				for (var i=0;i<Rec.Fields.Count;i++)
				{
					
						if (Rec.Fields.Item(i).Value==null)
							Result.appendChild(Rec.Fields.Item(i).Name);
						else
						{
							if ((Rec.Fields.Item(i).Type==129) && (Rec.Fields.Item(i).DefinedSize==28))
								Result.appendChild(Rec.Fields.Item(i).Name,Rec.Fields.Item(i).Value.replace(/[-]/g,"/").replace(/(^[ 　]+)|([ 　]+$)/g,""));
							else
								Result.appendChild(Rec.Fields.Item(i).Name,Rec.Fields.Item(i).Value.toString().replace(/(^[ 　]+)|([ 　]+$)/g,""));
						}
					
				}
				Result.toParent();
				Rec.MoveNext();
			}
			Rec.Close();
		}
	}
	
}catch(e){
	
	Result.appendChild("ERROR",e.description);
	
}finally{
	if (Conn.State==1)
	    	Conn.Close();
	CMD=null;
	Conn=null;
	Response.ContentType = "text/xml";
	Response.Write(Result.xml());
}


function Make_SQL_COMMAND(){

	var SQL_Select = ""+
		" select top "+MAX_QUERY_NUM+
		" sd.send_full_no , "+
		"am.send_no, "+
		"cd.SUBJECT, "+
		"Convert(char(28),sd.SENDER_TIME,20) as SENDER_TIME, "+
		"Convert(char(28),sd.SEND_DOC_DATE,20) as SEND_DOC_DATE, "+
//	" isnull(ap.acc_name,ap.ADBOOK_NAME) as ap_name , "+
		" isnull(ap.acc_name,ap.ADBOOK_NAME) as ap_name , "+
		" Convert(char(28),xsm.Confirm_Time,20) as Confirm_Time ,"+
		" case "+
			"when xsm.Confirm_Time is null then '0' "+
			"when xsm.Confirm_Time ='' then '0' "+
			"ELSE '1' "+
		"END as TMP_IS_SUCCESS ,"+
		"ap.id as AP_ID ,"+
		"xsm.id as XSM_ID ,"+
		"case "+
			"when xsm.is_user_confirm is null then '0' "+
			" when xsm.is_user_confirm = 0 then 0 "+
			" else '1' "+
		"end as IS_CONFIRM, "+
		"ap.acc_ORG_CODE, "+
		"ap.ACC_UNIT_CODE ,"+
		"CF.ID AS CFID , "+
		"CF.COMBINE_FLAG, "+
		"CF.FROM_SUBJECT, "+
		"CF.CREATE_SUBJECT, "+
		"CF.FLOW_SIGN_FLAG, "+
		"CF.CHARGE_USER_ID, "+
		"CF.CHARGE_DEPT_ID, "+
		"CF.CREATE_DOC_STYLE, "+
		"CF.DOC_FILE_ID, "+
		"CF.CREATE_FILE_ID, "+
		"CF.SEND_DOC_FLAG, "+
		"CF.COME_DOC_NO, "+
		"CF.RECV_DOC_NO, "+
		"CF.CREATE_DOC_NO, "+
		"CF.SPEED , "+
		"CF.UNDERTAKER_ACT, "+
		"xsm.User_note as DOING_REASON, "+
		"CF.CREATE_DOC_TYPE, "+
		"CF.SECRET, "+
		"CF.CHARGE_USER_ID AS CURRENT_USER_ID ";
			


	var SQL_FROM_1 = " from send_doc sd "+
			" inner join ACC_MAIN AM on sd.create_doc_id=am.doc_id "+
			" inner join create_doc	cd on cd.id=sd.create_doc_id "+
			" inner join CUR_FLOW	CF on cd.FLOW_ID=CF.ID "+
			" inner join accepter	ap on am.doc_id=ap.doc_id and am.DOCNAME=ap.DOCNAME"+
			" left outer join xmlbox_send_msg xsm on ap.acc_ORG_CODE+ap.ACC_UNIT_CODE=xsm.receiver "+
			" and sd.send_full_no=xsm.doc_no and am.send_no=xsm.document_id ";
			//" and xsm.confirm_time is not null ";
			
	var SQL_FROM_2 = " from send_doc sd "+
			" inner join ACC_MAIN AM on sd.create_doc_id=am.doc_id "+
			" inner join create_doc cd on cd.id=sd.create_doc_id "+
			" inner join HISTORY_FLOW	CF on cd.FLOW_ID=CF.ID "+
			" inner join accepter ap on am.doc_id=ap.doc_id and am.DOCNAME=ap.DOCNAME"+
			" left outer join xmlbox_send_msg xsm on ap.acc_ORG_CODE+ap.ACC_UNIT_CODE=xsm.receiver "+
			" and sd.send_full_no=xsm.doc_no  and am.send_no=xsm.document_id ";
			//" and xsm.confirm_time is not null ";
			
			
	var SQL_Where =" where am.send_no is not null and am.send_no !='' "+
		" and ap.deli_way=1 "+
		" and ap.acc_ORG_CODE is not null and ap.acc_ORG_CODE!='' ";

	
	
	
	//是否使用文號查詢
	if(doc_no_qry_flag == 1 )
	{	
		if (reqDoc.toChild("/REQUEST/SEND_NO") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			SQL_Where = SQL_Where + " and sd.send_full_no='"+reqDoc.getText()+"'";
		}	
	}
	else
	{
		var from_date="";
		var to_date="";
		
		if (reqDoc.toChild("/REQUEST/FROM_DATE") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			from_date=reqDoc.getText().replace("///g","-");
		}
		
		if (reqDoc.toChild("/REQUEST/TO_DATE") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			to_date=reqDoc.getText().replace("///g","-");
		}

			
		if(from_date!="" && to_date!=""){
			SQL_Where = SQL_Where + " and sd.SENDER_TIME between '"+from_date+"' and '"+to_date+"'";	
		}
		
		
		
		var from_date_send_doc="";
		var to_date_send_doc="";
		
		if (reqDoc.toChild("/REQUEST/FROM_DATE_SEND_DOC") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			from_date_send_doc=reqDoc.getText().replace("///g","-");
		}
		
		if (reqDoc.toChild("/REQUEST/TO_DATE_SEND_DOC") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			to_date_send_doc=reqDoc.getText().replace("///g","-");
		}
		
		if(from_date_send_doc!="" && to_date_send_doc!=""){
			SQL_Where = SQL_Where + " and sd.SEND_DOC_DATE between '"+from_date_send_doc+"' and '"+to_date_send_doc+"'";	
		}
		
		

		
		
		if (reqDoc.toChild("/REQUEST/QUERY_DEPT_ID") && reqDoc.getTextInt()!=null && reqDoc.getTextInt()!=0 ){
			SQL_Where = SQL_Where + " and sd.SEND_DEPT_ID='"+reqDoc.getTextInt()+"'";
		}
		
		
		
		//成功或失敗
		if (reqDoc.toChild("/REQUEST/ELEC_SEND_RESULT") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			if(reqDoc.getTextInt()==1){
				SQL_Where = SQL_Where + " and xsm.Confirm_Time is not null";	
			}else if(reqDoc.getTextInt()==2){
				SQL_Where = SQL_Where + " and xsm.Confirm_Time is null";
			}
		}
		//處理結果
		if (reqDoc.toChild("/REQUEST/DOING_RESULT") && reqDoc.getText()!=null && reqDoc.getText()!="" ){
			if(reqDoc.getTextInt()==1){
				SQL_Where = SQL_Where + " and xsm.Is_user_confirm =1";	
			}else if(reqDoc.getTextInt()==2){
				SQL_Where = SQL_Where + " and (xsm.Is_user_confirm =0 or xsm.Is_user_confirm is null)";	
			}
		}
	}//END ELSE (是否使用文號查詢)
	
	var QUERY_ORDER_COND=" order by am.send_no,ap.id ";
	

	var tmpSQL =  "("+ SQL_Select + SQL_FROM_1 + SQL_Where +") UNION ("+ SQL_Select + SQL_FROM_2 + SQL_Where +")"  + QUERY_ORDER_COND ;
	
	return tmpSQL;	
}
%>