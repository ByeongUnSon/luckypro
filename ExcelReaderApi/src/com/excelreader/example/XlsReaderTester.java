package com.excelreader.example;


import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.excelreader.dao.InputLayoutData;
import com.excelreader.util.XlsReader;

public class XlsReaderTester {
	static int temp = -1;
	static int cnt = 0;
	public static void main(String[] args) throws InvalidFormatException,
			IOException {

		Scanner sc = new Scanner(System.in);

		int sheetIdx = 0;
		int startSheetIdx = 0;
		int endSheetIdx = 0;
		int startRowIdx = 0;
		int endRowIdx = 0;
		int startColIdx = 0;
		int endColIdx = 0;
		List<String> layoutDataList = null;
		// [filePath]�� test.xls ���Ϸ� �׽�Ʈ Ư����Ʈ 2���̻� 
		if (args[0].length() == 0) {
			System.err.println("java XlsReaderTester [filePath]");
		} else {

			XlsReader reader = new XlsReader(args[0]);

			reader.start();
							
			int select = 0;
			
			do {
				System.out.println("\nExcel Layout Extract Program");
				System.out.println("1. Ư�� ��Ʈ, ���� ��, �� ��, ���� ��, �� ��");  // 2, 9, 46, 3, 5 �� �׽�Ʈ
				System.out.println("2. ���� ��Ʈ, �� ��Ʈ, ���� ��, �� ��, ���� ��, �� ��"); // 2, 6, 9, 46, 3, 5 �� �׽�Ʈ
				System.out.print("��ȣ ���� : ");
				select = sc.nextInt();

				switch (select) {
				case 1:
					System.out.print("Ư�� ��Ʈ : ");
					sheetIdx = sc.nextInt();
					System.out.print("���� �� : ");
					startRowIdx = sc.nextInt();
					System.out.print("�� �� : ");
					endRowIdx = sc.nextInt();
					System.out.print("���� �� : ");
					startColIdx = sc.nextInt();
					System.out.print("�� �� : ");
					endColIdx = sc.nextInt();
					layoutDataList = reader.getRowCellsAtStartRowColIdx(
							sheetIdx, startRowIdx, endRowIdx, startColIdx,
							endColIdx);
					
		
					LinkedList<InputLayoutData> inputData = new LinkedList<InputLayoutData>();

					InputLayoutData data = new InputLayoutData();
					
					int totalSize = layoutDataList.size();
				
				
					while (cnt <= totalSize - 1) {		
						temp++;
						switch (temp) {
							case 0:
								data.setDataOne(layoutDataList.get(cnt));	
								cnt++;
								break;
							case 1:
								data.setDataTwo(layoutDataList.get(cnt));
								cnt++;
								break;
							case 2:		
								data.setDataThree(layoutDataList.get(cnt));	
								inputData.add(data);
								data = new InputLayoutData();
								cnt++;
								temp = -1;
								break;
						}						
					}
					

					System.out.println();
					
					for (int i=0; i<inputData.size(); i++) {
						System.out.print("[ ");
						System.out.print(inputData.get(i).getDataOne() + ", ");
						System.out.print(inputData.get(i).getDataTwo() + ", ");
						System.out.println(inputData.get(i).getDataThree() + "   ] ");
					}
					
					//printResult(layoutDataList, startColIdx, endColIdx);
					//dao.saveLayoutData(layoutDataList, startColIdx, endColIdx);
					//System.out.println(dao.getLayoutDataList());
				
					//printResult(dao.getLayoutDataList(), startColIdx, endColIdx);
					break;
				case 2:
					System.out.print("���� ��Ʈ : ");
					startSheetIdx = sc.nextInt();
					System.out.print("�� ��Ʈ : ");
					endSheetIdx = sc.nextInt();
					System.out.print("���� �� : ");
					startRowIdx = sc.nextInt();
					System.out.print("�� �� : ");
					endRowIdx = sc.nextInt();
					System.out.print("���� �� : ");
					startColIdx = sc.nextInt();
					System.out.print("�� �� : ");
					endColIdx = sc.nextInt();
					layoutDataList = reader.getRowCellsAtStartRowColIdx(
							startSheetIdx, endSheetIdx, startRowIdx, endRowIdx,
							startColIdx, endColIdx);

					inputData = new LinkedList<InputLayoutData>();

					data = new InputLayoutData();
					
					totalSize = layoutDataList.size();
				
				
					while (cnt <= totalSize - 1) {		
						temp++;
						switch (temp) {
							case 0:
								data.setDataOne(layoutDataList.get(cnt));	
								cnt++;
								break;
							case 1:
								data.setDataTwo(layoutDataList.get(cnt));
								cnt++;
								break;
							case 2:		
								data.setDataThree(layoutDataList.get(cnt));	
								inputData.add(data);
								data = new InputLayoutData();
								cnt++;
								temp = -1;
								break;
						}						
					}
					

					System.out.println();
					
					for (int i=0; i<inputData.size(); i++) {
						System.out.print("[ ");
						System.out.print(inputData.get(i).getDataOne() + ", ");
						System.out.print(inputData.get(i).getDataTwo() + ", ");
						System.out.println(inputData.get(i).getDataThree() + "   ] ");
					}
					break;
					default:
				}

			} while (select == 1 || select == 2);

		
		}
			

	}
	

	

}
