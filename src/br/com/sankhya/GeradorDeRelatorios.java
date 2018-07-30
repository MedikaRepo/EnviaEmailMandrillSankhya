package br.com.sankhya;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;
import net.sf.jasperreports.engine.JRExporter;
import net.sf.jasperreports.engine.JRExporterParameter;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.export.JRPdfExporter;

public class GeradorDeRelatorios
{ 
	public static void geraPdf(String jrxml, BigDecimal pNunota) 
	{
		try 
		{
			@SuppressWarnings({ "unchecked", "rawtypes" })
			Map<String, Object> parametros = new HashMap();
			parametros.put("NUNOTA", pNunota);
			//OutputStream saida = new FileOutputStream("/users/adriano/relatorios/propvenda/propvenda.pdf");
			OutputStream saida = new FileOutputStream("/home/mgeweb/modelos/relatorios/propostadevenda/propvenda.pdf");
			//OutputStream saida = new FileOutputStream("C:\\Users\\STI-004\\propostadevenda\\propvenda.pdf");

			// compila jrxml em memoria
			try {
				JasperReport jasper = JasperCompileManager.compileReport(jrxml);

				// preenche relatorio
				JasperPrint print = JasperFillManager.fillReport(jasper, parametros, ConnectMSSQLServer.conn);

				// exporta para pdf
				JRExporter exporter = new JRPdfExporter();
				exporter.setParameter(JRExporterParameter.JASPER_PRINT, print);
				exporter.setParameter(JRExporterParameter.OUTPUT_STREAM, saida);

				exporter.exportReport();

			}
			catch(Exception e){
				//System.out.println(e.getMessage());
				throw new RuntimeException("Erro ao gerar relatório:"+e.getMessage(), e);
			}

		}
		catch (Exception e) 
		{
			throw new RuntimeException("Erro ao gerar relatório"+e.getMessage(), e);
		}
	}   
}
