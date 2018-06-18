package br.com.sankhya;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.commons.codec.binary.Base64;
import com.microtripit.mandrillapp.lutung.MandrillApi;
import com.microtripit.mandrillapp.lutung.model.MandrillApiError;
import com.microtripit.mandrillapp.lutung.view.MandrillMessage;
import com.microtripit.mandrillapp.lutung.view.MandrillMessage.MessageContent;
import com.microtripit.mandrillapp.lutung.view.MandrillMessage.Recipient;
import com.microtripit.mandrillapp.lutung.view.MandrillMessage.Recipient.Type;
import com.microtripit.mandrillapp.lutung.view.MandrillMessageStatus;
import br.com.sankhya.extensions.actionbutton.AcaoRotinaJava;
import br.com.sankhya.extensions.actionbutton.ContextoAcao;
import br.com.sankhya.extensions.actionbutton.Registro;

//Atualizacao do catalogo - GAS - 18062018

public class SendMsg implements AcaoRotinaJava 

{
	static StringBuffer mensagem = new StringBuffer();

	private String getDateTime() 
	{ 
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss"); 
		Date date = new Date(); 
		return dateFormat.format(date); 
	}

	public static Object ExecutaComandoNoBanco(String sql, String op, String exibe)
	{
		try
		{
			Statement smnt = ConnectMSSQLServer.conn.createStatement(); 

			if(op=="select")
			{

				smnt.execute(sql);

				ResultSet result = smnt.executeQuery(sql);

				result.next();

				return result.getObject(1);
			}
			else if(op=="alter")
			{
				smnt.executeUpdate(sql);
				return (Object)1;
			}
			else
			{
				return null;
			}
		}
		catch(SQLException ex)
		{
			System.err.println("SQLException: " + ex.getMessage());
			mensagem.append("Erro ao obter campo SQL("+ex.getMessage()+") \n");
			return null;
		}
	}	

	//public static void main(String[] args) throws IOException 
	@SuppressWarnings({ "resource", "unused" })
	@Override
	public void doAction(ContextoAcao contexto) throws Exception 

	{
		String emailParc="", nomeVend="", emailVend="", textoAd="", emailCopia="", emailCopia2="", emailCopia3="", nomeParc="", emailLider="", nomeLider="";
		BigDecimal nunota=BigDecimal.ZERO, tipoOp=BigDecimal.ZERO, codParc=BigDecimal.ZERO, 
				codVend=BigDecimal.ZERO, codContato=BigDecimal.ZERO;
		int ramalVend=0;
		MandrillApi mandrillApi = new MandrillApi("5ZbB4mED1LcYaISfhPtquQ");

		mensagem.setLength(0);

		//Conecta no banco do Sankhya
		ConnectMSSQLServer.dbConnect("jdbc:sqlserver://192.168.0.5:1433;DatabaseName=SANKHYA_PROD;", "adriano","Compiles23");

		//recupera o numero da negociação
		try
		{
			Registro[] linha = contexto.getLinhas();

			for (int i = 0; i < linha.length; i++) 
			{

				nunota= (BigDecimal)linha[i].getCampo("NUNOTA");
				tipoOp=(BigDecimal) linha[i].getCampo("CODTIPOPER");
				codParc=(BigDecimal) linha[i].getCampo("CODPARC");
				codVend=(BigDecimal) linha[i].getCampo("CODVEND");
				codContato=(BigDecimal) linha[i].getCampo("CODCONTATO");
				emailCopia=(String)linha[i].getCampo("AD_COPIAEMAILPROP");
				emailCopia2=(String)linha[i].getCampo("AD_COPIA2EMAILPROP");
				emailCopia3=(String)linha[i].getCampo("AD_COPIA3EMAILPROP");
			}
		}catch(Exception e)
		{
			mensagem.append("Erro ao obter campos sankhya. "+e.getMessage());
			contexto.setMensagemRetorno(mensagem.toString());
		}


		//Só executa nas TOPs permitidas   
		if ((tipoOp.intValue()==204)||(tipoOp.intValue()==205))
		{	
			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT AD_TEXTOADICIONALEMAILPROP FROM TGFCAB WHERE NUNOTA="
					+nunota.toString(), "select", "N")!=null)
			{
				textoAd=(String) ExecutaComandoNoBanco("SELECT AD_TEXTOADICIONALEMAILPROP FROM TGFCAB WHERE NUNOTA="
						+nunota.toString(), "select", "N");
			}

			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT NOMEPARC FROM TGFPAR WHERE CODPARC="
					+codParc.toString(), "select", "N")!=null)
			{
				nomeParc=(String) ExecutaComandoNoBanco("SELECT NOMEPARC FROM TGFPAR WHERE CODPARC="
						+codParc.toString(), "select", "N");
			}

			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU WHERE CODUSU="
					+contexto.getUsuarioLogado().toString(), "select", "N")!=null)
			{
				ramalVend=(Integer) ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU WHERE CODUSU="
						+contexto.getUsuarioLogado().toString(), "select", "N");
			}

			//Recupera o email do vendedor
			if(ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
					+contexto.getUsuarioLogado().toString(), "select", "N")!=null)
			{
				emailVend=(String) ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
						+contexto.getUsuarioLogado().toString(), "select", "N");
			}

			//Recupera o nome do vendedor
			if(ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
					" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
					" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
					" WHERE  USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select", "N")!=null)
			{
				nomeVend=(String) ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
						" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
						" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
						" WHERE USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select", "N");
			}

			//Recupera o email do parceiro
			if(ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM TGFCAB CAB"
					+ " INNER JOIN TGFCTT CTT ON CTT.CODCONTATO=CAB.CODCONTATO"
					+ " AND CTT.CODPARC=CAB.CODPARC"
					+ " WHERE CAB.CODCONTATO="+codContato.toString()+" AND CAB.CODPARC="
					+codParc.toString(), "select", "S")!=null)
			{
				emailParc=(String) ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM TGFCAB CAB"
						+ " INNER JOIN TGFCTT CTT ON CTT.CODCONTATO=CAB.CODCONTATO"
						+ " AND CTT.CODPARC=CAB.CODPARC"
						+ " WHERE CAB.CODCONTATO="+codContato.toString()+" AND CAB.CODPARC="
						+codParc.toString(), "select", "N");
				
			}
			
			//Recupera o email da lider
			if(ExecutaComandoNoBanco("SELECT USU2.EMAIL FROM TSIUSU USU "+
					" INNER JOIN TSIUSU USU2 on USU2.CODUSU = USU.AD_CODUSUSUP where USU.CODUSU="
					+contexto.getUsuarioLogado().toString(), "select", "S")!=null)
			{
				emailLider=(String) ExecutaComandoNoBanco("SELECT USU2.EMAIL FROM TSIUSU USU "+
						" INNER JOIN TSIUSU USU2 on USU2.CODUSU = USU.AD_CODUSUSUP where USU.CODUSU="
						+contexto.getUsuarioLogado().toString(), "select", "N");
			}

			//Recupera o nome da lider
			if(ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN "+
					"INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND "+
					"INNER JOIN TSIUSU USU2 ON USU2.CODUSU = USU.AD_CODUSUSUP "+
					"INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU2.CODFUNC AND FUN.CODEMP=3 "+
					" WHERE  USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select", "N")!=null)
			{
				nomeLider=(String) ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN "+
						"INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND "+
						"INNER JOIN TSIUSU USU2 ON USU2.CODUSU = USU.AD_CODUSUSUP "+
						"INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU2.CODFUNC AND FUN.CODEMP=3 "+
						" WHERE  USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select", "N");
			}
			
			contexto.setMensagemRetorno(mensagem.toString());

			try
			{
				// create your message
				MandrillMessage message = new MandrillMessage();

				message.setSubject("Medika - Proposta Comercial - "+nunota.toString());
				if (textoAd=="")
				{
					message.setHtml("<html><body style="+"\"font-famaly: arial; font-size:12px;"+"\">Prezado(s),<br/><br/>"+		             
							"Segue proposta(s) de fornecimento do(s) material(ais) importado(s) e comercializado(s) pela Medika.<br/><br/>"+
							"Caso deseje visualizar nossos cat�logos de produtos clique neste link: <a href="+
							"\"https://www.medika.com.br/catalogo\">Cat�logos Medika</a>"+
							"<br/><br/>"+
							"Atenciosamente,"+
							"<br/><br/>"+nomeVend+
							" - Tel:(31) 3688-1901 Ramal:"+ramalVend+" - Equipe de Vendas"+
							"<br><br><HR WIDTH=100% style="+"\"border:1px solid #191970;"+
							"\"><img src="+"\"https://static.wixstatic.com/media/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png/v1/fill/w_251,h_104,al_c,usm_0.66_1.00_0.01/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png"+
							"\"><br><br>Medika, qualidade em sa�de. - <a href="+"\"http://www.medika.com.br"+
							"\">www.medika.com.br</a><br>"+
							"<HR WIDTH=100% style="+"\"border:1px solid #191970;"+"\">"+
							"</body></html>");
				}
				else
				{
					message.setHtml("<html><body style="+"\"font-famaly: arial; font-size:12px;"+"\">Prezado(s),<br/><br/>"+		             
							"Segue proposta(s) de fornecimento do(s) material(ais) importado(s) e comercializado(s) pela Medika.<br/><br/>"+
							"<br/>"+textoAd+"<br/><br/><br/>"
							+ "Caso deseje visualizar nossos cat�logos de produtos clique neste link: <a href="+
							"\"https://www.medika.com.br/catalogo\">Cat�logos Medika</a>"+
							"<br/><br/>"+
							"Atenciosamente,"+
							"<br/><br/>"+nomeVend+
							" - Tel:(31) 3688-1901 Ramal:"+ramalVend+" - Equipe de Vendas"+
							"<br><br><HR WIDTH=100% style="+"\"border:1px solid #191970;"+
							"\"><img src="+"\"https://static.wixstatic.com/media/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png/v1/fill/w_251,h_104,al_c,usm_0.66_1.00_0.01/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png"+
							"\"><br><br>Medika, qualidade em sa�de. - <a href="+"\"http://www.medika.com.br"+
							"\">www.medika.com.br</a><br>"+
							"<HR WIDTH=100% style="+"\"border:1px solid #191970;"+"\">"+
							"</body></html>");	
				}
				message.setAutoText(true);
				message.setFromEmail(emailVend);
				message.setFromName("Equipe de Vendas - Medika");

				// add recipients
				ArrayList<Recipient> recipients = new ArrayList<Recipient>();
				Recipient recipient = new Recipient();
				recipient.setEmail(emailParc);
				recipient.setName(nomeParc);
				recipient.setType(Type.TO);
				recipients.add(recipient);

				//EMAIL CÓPIA
				Recipient recipient2 = new Recipient();
				recipient2.setEmail(emailCopia);
				recipient2.setType(Type.BCC);
				recipients.add(recipient2);

				//EMAIL CÓPIA 2
				Recipient recipient3 = new Recipient();
				recipient3.setEmail(emailCopia2);
				recipient3.setType(Type.BCC);
				recipients.add(recipient3);
				
				//EMAIL CÓPIA 3
				Recipient recipient4 = new Recipient();
				recipient4.setEmail(emailCopia3);
				recipient4.setType(Type.BCC);
				recipients.add(recipient4);
				
				//EMAIL CÓPIA Email Vendedor 
				Recipient recipient5 = new Recipient();
				recipient5.setEmail(emailVend);
				recipient5.setName(nomeVend);
				recipient5.setType(Type.BCC);
				recipients.add(recipient5);
				
				//EMAIL CÓPIA Email Vendedor 
				Recipient recipient6 = new Recipient();
				recipient6.setEmail(emailLider);
				recipient6.setName(nomeLider);
				recipient6.setType(Type.BCC);
				recipients.add(recipient6);

				message.setTo(recipients);
				message.setPreserveRecipients(true);
				message.setTrackOpens(true);
				message.setTrackClicks(true);

				//Tratamento para adicionar anexo na mensagem
				List<MessageContent> listofAttachments = new ArrayList<MessageContent>();
				MessageContent attachment = new MessageContent();
				attachment.setType("application/pdf");
				attachment.setName("Proposta Comercial.pdf");

				//Criação do anexo da proposta
				GeradorDeRelatorios.geraPdf("/home/mgeweb/modelos/relatorios/propostadevenda/PEDIDO_DE_VENDA1x.jrxml", nunota);  
				//GeradorDeRelatorios.geraPdf("/users/adriano/relatorios/propvenda/PEDIDO_DE_VENDA1x.jrxml", nunota.add(new BigDecimal(122814)));

				//File file = new File("/users/adriano/relatorios/propvenda/propvenda.pdf");
				File file = new File("/home/mgeweb/modelos/relatorios/propostadevenda/propvenda.pdf");

				InputStream is = new FileInputStream(file);

				long length = file.length();
				if (length > Integer.MAX_VALUE) 
				{
					// File is too large
				}
				byte[] bytes = new byte[(int) length];

				int offset = 0;
				int numRead = 0;
				while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) 
				{
					offset += numRead;
				}

				if (offset < bytes.length) 
				{
					throw new IOException("Não foi possível completar a leitura do arquivo. " + file.getName());
				}

				is.close();
				byte[] encoded = Base64.encodeBase64(bytes);
				String encodedString = new String(encoded);
				attachment.setContent(encodedString);
				listofAttachments.add(attachment);      

				message.setAttachments(listofAttachments);

				int proxCodEmail;

				ArrayList<String> tags = new ArrayList<String>();
				tags.add(nunota.toString());
				proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select", "N");
				tags.add(Integer.toString(proxCodEmail+1));

				Calendar cal = Calendar.getInstance();
		        
		        int month = cal.get(Calendar.MONTH);
		        int day = cal.get(Calendar.DAY_OF_MONTH);
		        int year = cal.get(Calendar.YEAR);
		        int hour = cal.get(Calendar.HOUR_OF_DAY);
		        int minute = cal.get(Calendar.MINUTE);
		        int second = cal.get(Calendar.SECOND);
		        
		        String dhenvio = day +"/"+month+1+"/"+year+" "+hour+":"+minute+":"+second;
				
				
				if(ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', getdate(), '"+ emailParc+"','"+emailVend+"')", "alter", "N")!=null)
		        //if(ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailParc+"','"+emailVend+"')", "alter")!=null)
		        {
					mensagem.append("Registros adicionados na tabela de log de envios.\n\n");
				}

				if ((emailCopia!="null")&&(emailCopia!=null))
				{
					proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select", "N");
					ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', getdate(),'"+ emailCopia+"','"+emailVend+"')", "alter", "N");
					//tags.add(Integer.toString(proxCodEmail+1));
				}
				if ((emailCopia2!="null")&&(emailCopia2!=null))
				{
					proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select", "N");
					ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', getdate(),'"+ emailCopia2+"','"+emailVend+"')", "alter", "N");
					//tags.add(Integer.toString(proxCodEmail+1));
				}
				message.setTags(tags);

				// ... add more message details if you want to!
				// then ... send
				try {
					MandrillMessageStatus[] messageStatusReports = mandrillApi
							.messages().send(message, false);

					mensagem.append("Email enviado com sucesso! \n");

					contexto.setMensagemRetorno(mensagem.toString());
					//System.out.println(mensagem);

				} catch (MandrillApiError e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					//mensagem.setLength(0);
					mensagem.append("Email do Contato:" + emailParc + "\n Email do Vendedor:" + emailVend + "\n Email do Líder:" + emailLider + "\n");
					mensagem.append("Contato selecionado não tem email cadastrado ou o vendedor não tem email cadastrado. Favor verificar. "+e.getMessage()+"\n");
					mensagem.append("Erro na API do MailChimp. "+e.getMessage()+"\n");
					contexto.setMensagemRetorno(mensagem.toString());
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					mensagem.append("Erro de IO, ao tentar anexar pdf ao email do MailChimp. "+e.getMessage()+"\n");
					contexto.setMensagemRetorno(mensagem.toString());
				}
			}catch(Exception e)
			{
				mensagem.append("Erro ao enviar mensagem pelo MailChimp. "+e.getMessage()+"\n");
				contexto.setMensagemRetorno(mensagem.toString());
			}
		}
		else
		{
			//StringBuffer mensagem = new StringBuffer();
			mensagem.append("O envio da Proposta de Venda e Apresentação Comercial deve ser feito apenas nas TOPs 204 ou 205!");

			contexto.setMensagemRetorno(mensagem.toString());
			//System.out.println("O envio da Proposta de Venda e Apresentação Comercial deve ser feito apenas na TOP 204");  
		}

	}

}
