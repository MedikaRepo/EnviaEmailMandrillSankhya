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

public class SendMsg implements AcaoRotinaJava 

{
	static StringBuffer mensagem = new StringBuffer();

	private String getDateTime() 
	{ 
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss"); 
		Date date = new Date(); 
		return dateFormat.format(date); 
	}

	public static Object ExecutaComandoNoBanco(String sql, String op)
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
		String emailParc="", nomeVend="", emailVend="", textoAd="", emailCopia="", emailCopiaOculta="", nomeParc="";
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
				emailCopiaOculta=(String)linha[i].getCampo("AD_COPIA2EMAILPROP");
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
					+nunota.toString(), "select")!=null)
			{
				textoAd=(String) ExecutaComandoNoBanco("SELECT AD_TEXTOADICIONALEMAILPROP FROM TGFCAB WHERE NUNOTA="
						+nunota.toString(), "select");
			}

			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT NOMEPARC FROM TGFPAR WHERE CODPARC="
					+codParc.toString(), "select")!=null)
			{
				nomeParc=(String) ExecutaComandoNoBanco("SELECT NOMEPARC FROM TGFPAR WHERE CODPARC="
						+codParc.toString(), "select");
			}

			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU WHERE CODUSU="
					+contexto.getUsuarioLogado().toString(), "select")!=null)
			{
				ramalVend=(Integer) ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU WHERE CODUSU="
						+contexto.getUsuarioLogado().toString(), "select");
			}

			//Recupera o email do vendedor
			if(ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
					+contexto.getUsuarioLogado().toString(), "select")!=null)
			{
				emailVend=(String) ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
						+contexto.getUsuarioLogado().toString(), "select");
			}

			//Recupera o nome do vendedor
			if(ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
					" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
					" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
					" WHERE  USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select")!=null)
			{
				nomeVend=(String) ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
						" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
						" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
						" WHERE USU.DTLIMACESSO IS NULL AND FUN.DTDEM IS NULL AND USU.CODVEND="+codVend.toString(), "select");
			}

			//Recupera o email do parceiro
			if(ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM TGFCAB CAB"
					+ " INNER JOIN TGFCTT CTT ON CTT.CODCONTATO=CAB.CODCONTATO"
					+ " AND CTT.CODPARC=CAB.CODPARC"
					+ " WHERE CAB.CODCONTATO="+codContato.toString()+" AND CAB.CODPARC="
					+codParc.toString(), "select")!=null)
			{
				emailParc=(String) ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM TGFCAB CAB"
						+ " INNER JOIN TGFCTT CTT ON CTT.CODCONTATO=CAB.CODCONTATO"
						+ " AND CTT.CODPARC=CAB.CODPARC"
						+ " WHERE CAB.CODCONTATO="+codContato.toString()+" AND CAB.CODPARC="
						+codParc.toString(), "select");
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
							"Caso deseje visualizar nossos catálogos de produtos clique neste link: <a href="+
							"\"https://1drv.ms/f/s!AhNxsawi2mh1q1G2xpw7_AtidD3m\">Catálogos Medika</a>"+
							"<br/><br/>"+
							"Atenciosamente,"+
							"<br/><br/>"+nomeVend+
							" - Tel:(31) 3688-1901 Ramal:"+ramalVend+" - Equipe de Vendas"+
							"<br><br><HR WIDTH=100% style="+"\"border:1px solid #191970;"+
							"\"><img src="+"\"http://www.medika.com.br/wp-content/uploads/2016/05/logo-medika.png"+
							"\"><br><br>Medika, qualidade em saúde. - <a href="+"\"http://www.medika.com.br"+
							"\">www.medika.com.br</a><br>"+
							"<HR WIDTH=100% style="+"\"border:1px solid #191970;"+"\">"+
							"</body></html>");
				}
				else
				{
					message.setHtml("<html><body style="+"\"font-famaly: arial; font-size:12px;"+"\">Prezado(s),<br/><br/>"+		             
							"Segue proposta(s) de fornecimento do(s) material(ais) importado(s) e comercializado(s) pela Medika.<br/><br/>"+
							"<br/>"+textoAd+"<br/><br/><br/>"
							+ "Caso deseje visualizar nossos catálogos de produtos clique neste link: <a href="+
							"\"https://www.medika.com.br/catalogo</a>"+
							"<br/><br/>"+
							"Atenciosamente,"+
							"<br/><br/>"+nomeVend+
							" - Tel:(31) 3688-1901 Ramal:"+ramalVend+" - Equipe de Vendas"+
							"<br><br><HR WIDTH=100% style="+"\"border:1px solid #191970;"+
							"\"><img src="+"\"http://www.medika.com.br/wp-content/uploads/2016/05/logo-medika.png"+
							"\"><br><br>Medika, qualidade em saúde. - <a href="+"\"http://www.medika.com.br"+
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
				recipient2.setName(nomeParc);
				recipient2.setType(Type.CC);
				recipients.add(recipient2);

				//EMAIL CÓPIA 2
				Recipient recipient3 = new Recipient();
				recipient3.setEmail(emailCopiaOculta);
				recipient3.setName(nomeParc);
				recipient3.setType(Type.CC);
				recipients.add(recipient3);

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
				proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select");
				tags.add(Integer.toString(proxCodEmail+1));

				if(ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailParc+"','"+emailVend+"')", "alter")!=null)
				{
					mensagem.append("Registros adicionados na tabela de log de envios.\n\n");
				}

				if ((emailCopia!="null")&&(emailCopia!=null))
				{
					proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select");
					ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailCopia+"','"+emailVend+"')", "alter");
					tags.add(Integer.toString(proxCodEmail+1));
				}
				if ((emailCopiaOculta!="null")&&(emailCopiaOculta!=null))
				{
					proxCodEmail=(Integer)ExecutaComandoNoBanco("SELECT MAX(CODEMAILMONITOR) FROM AD_EMAILMONITOR", "select");
					ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailCopiaOculta+"','"+emailVend+"')", "alter");
					tags.add(Integer.toString(proxCodEmail+1));
				}
				message.setTags(tags);

				// ... add more message details if you want to!
				// then ... send
				try {
					MandrillMessageStatus[] messageStatusReports = mandrillApi
							.messages().send(message, false);

					mensagem.append("Email enviado com sucesso! \n");

					//Insere log de envio no banco
					/*try
					{
						if(emailCopia!=null)
						{
							ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailCopia+"','"+emailVend+"')", "alter");
						}

						if(emailCopiaOculta!=null)
						{
							ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailCopiaOculta+"','"+emailVend+"')", "alter");
						}

						if(ExecutaComandoNoBanco("INSERT INTO AD_EMAILMONITOR (NUNOTA, STATUSENVIO, DHENVIO, DESTINATARIO, REMETENTE)VALUES("+nunota.toString()+",'ENVIADO', '"+getDateTime()+"','"+ emailParc+"','"+emailVend+"')", "alter")!=null)
						{
							mensagem.append("Registros adicionados na tabela de log de envios.");
						}

					}
					catch (Exception e) 
					{
						mensagem.append("Erro ao inserir na registro na tabela de log de envios. "+e.getMessage());
						contexto.setMensagemRetorno(mensagem.toString());
					}*/

					contexto.setMensagemRetorno(mensagem.toString());
					//System.out.println(mensagem);

				} catch (MandrillApiError e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
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
