package main;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;



public class MainClass {

	public static void main(String[] args) throws Exception {
		
		String caminho = null;
		
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		
		int returnValue = jfc.showOpenDialog(null);
		
		if (returnValue== JFileChooser.APPROVE_OPTION) {
				
			File selectedFile = jfc.getSelectedFile();
			caminho = selectedFile.getAbsolutePath();
			System.out.println(selectedFile.getAbsolutePath());
			System.out.println(caminho);			
			
		}
		
		String dataDir = caminho;	
				
		// Abrindo arquivos CSV
		/// Criando o objeto CSV LoadOptions
		LoadOptions loadOptions = new LoadOptions(FileFormatType.CSV);		
		// Criando um objeto Workbook com caminho de arquivo CSV e loadOptions
		// objeto
		Workbook workbook = new Workbook(dataDir, loadOptions);
		workbook.save(dataDir + "CSVtoExcel.xlsx", SaveFormat.XLSX);
		System.out.println("Done");
		

	}

}
