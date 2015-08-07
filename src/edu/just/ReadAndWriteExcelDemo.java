package edu.just;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.omg.CORBA.PUBLIC_MEMBER;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
/**
 * 
 * @author Panshunxing
 *
 */
public class ReadAndWriteExcelDemo {
	/**
	 * �õ�workbook����
	 * @param filename 
	 * @return
	 */
	public Workbook getWorkBook(String filename){
		InputStream is = null;
		Workbook readbook=null;
		try {
			is = new FileInputStream(filename);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			readbook=Workbook.getWorkbook(is);
		} catch (BiffException | IOException e) {
			e.printStackTrace();
		}
		return readbook;
	}
	/**
	 * չʾȡ��excel������ݽ��
	 * @param wb
	 */
	public void showResult(Workbook wb){
		//��ȡ��һ��sheet��
		Sheet sheet=wb.getSheet(0);
		//��ȡ��һ��Sheet�������
		int numOfRow=sheet.getRows();
		//��ȡ��һ��Sheet�������
		int numOfColumn=sheet.getColumns();
		for(int i=0;i<numOfRow;i++){
			for(int j=0;j<numOfColumn;j++){
				Cell cell=sheet.getCell(j, i);
				System.out.print(cell.getContents()+"\t");
			}
			System.out.println();
		}
	}
	/**
	 * ������д�뵽Excel����
	 * @param cell
	 * @param file
	 */
	public void writeToWorkBook(Label cell,WritableSheet sheet){
		try {
			sheet.addCell(cell);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
	}
	public static void main(String[] args) {
		ReadAndWriteExcelDemo readExcelDemo=new ReadAndWriteExcelDemo();
		Workbook wb=readExcelDemo.getWorkBook("File\\JA1015-����+���Ǳ�γ̱�.xls");	
		readExcelDemo.showResult(wb);
		int[][] contents=new int[10][10];
		for(int i=0;i<contents.length;i++){
			for(int j=0;j<contents[0].length;j++){
				contents[i][j]=i*j;
			}
		}
		File file=new File("File\\1.xls");
		//����������
		WritableWorkbook workbook=null;
		try {
			workbook=Workbook.createWorkbook(file);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//����sheetҳ
		WritableSheet sheet=workbook.createSheet("FirstSheet", 0);
		for(int i=0;i<contents.length;i++){
			for(int j=0;j<contents[0].length;j++){
				Label label=new Label(i, j, contents[i][j]+"");
				readExcelDemo.writeToWorkBook(label, sheet);
			}
		}
		try {
			workbook.write();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			workbook.close();
		} catch (WriteException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
