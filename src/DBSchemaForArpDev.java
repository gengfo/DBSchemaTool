

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * Add your class description here...
 * 
 * @author GENGFO
 * @version 1.0, 2006-6-21
 */
public class DBSchemaForArpDev {

	final static String driverClass = "oracle.jdbc.driver.OracleDriver";

	final static String connectionURL = "jdbc:oracle:thin:@146.222.65.142:1521:irarpdb";

	final static String userID = "arpowner";

	final static String userPassword = "arpowner";

	Connection con = null;

	public DBSchemaForArpDev() {

		try {

			System.out.print("  Loading JDBC Driver  -> " + driverClass + "\n");
			Class.forName(driverClass).newInstance();

			System.out.print("  Connecting to        -> " + connectionURL
					+ "\n");
			this.con = DriverManager.getConnection(connectionURL, userID,
					userPassword);
			System.out.print("  Connected as         -> " + userID + "\n");

		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}

	}

	private static void prt(Object s) {
		System.out.print(s);
	}

	private static void prt(int i) {
		System.out.print(i);
	}

	private static void prtln(Object s) {
		prt(s + "\n");
	}

	private static void prtln(int i) {
		prt(i + "\n");
	}

	private static void prtln() {
		prtln("");
	}

	/**
	 * Close down Oracle connection.
	 */
	public void closeConnection() {

		try {
			System.out.print("  Closing Connection...\n");
			con.close();

		} catch (SQLException e) {

			e.printStackTrace();

		}

	}

	public void writeTablesToExcel() throws SQLException, BiffException,
			IOException, RowsExceededException, WriteException {

		DatabaseMetaData md = null;
		md = con.getMetaData();

		List tableList = new ArrayList();
		tableList = readTables();

		WritableWorkbook workbook = Workbook
				.createWorkbook(new File(
						"ARPDBSchema.xls"));
		WritableSheet sheet = workbook.createSheet("All Tables", 0);

		// show all tables
		for (int i = 0; i < tableList.size(); i++) {
			Label label = new Label(0, i + 1, (String) tableList.get(i));
			sheet.addCell(label);
		}

		//md.getIndexInfo(catalog, schema, table, unique, approximate);
		
		// show specified table
		for (int j = 0; j < tableList.size(); j++) {
			String tableName = (String) tableList.get(j);
			if (tableName.startsWith("BIN")) {
				continue;
			}
			WritableSheet tabelSheet = workbook.createSheet(tableName, 0);

			System.out.println("the table name is: " + tableName);
			ResultSet tableSchemas = md.getColumns("", "%", tableName, "%");

			Label labelCOLUMNNAME = new Label(1, 0, "COLUMN_NAME");
			Label labelDATA_TYPE = new Label(2, 0, "DATA_TYPE");
			Label labelTYPE_NAME = new Label(3, 0, "TYPE_NAME");
			Label labelCOLUMN_SIZE = new Label(4, 0, "COLUMN_SIZE");
			Label labelNULLABLE = new Label(5, 0, "NULLABLE");
			Label labelREMARKS = new Label(6, 0, "REMARKS");

			tabelSheet.addCell(labelCOLUMNNAME);
			tabelSheet.addCell(labelDATA_TYPE);
			tabelSheet.addCell(labelTYPE_NAME);
			tabelSheet.addCell(labelCOLUMN_SIZE);
			tabelSheet.addCell(labelNULLABLE);
			tabelSheet.addCell(labelREMARKS);

			int k = 0;
			while (tableSchemas.next()) {
				System.out.print("COLUMN_NAME = ");
				System.out.print(tableSchemas.getString("COLUMN_NAME"));
				System.out.print(" TYPE_NAME = ");
				System.out.println(tableSchemas.getString("TYPE_NAME"));

				Label l1 = new Label(1, k + 1, tableSchemas
						.getString("COLUMN_NAME"));
				Label l2 = new Label(2, k + 1, tableSchemas
						.getString("DATA_TYPE"));
				Label l3 = new Label(3, k + 1, tableSchemas
						.getString("TYPE_NAME"));
				Label l4 = new Label(4, k + 1, tableSchemas
						.getString("COLUMN_SIZE"));
				Label l5 = new Label(5, k + 1, tableSchemas
						.getString("NULLABLE"));
				Label l6 = new Label(6, k + 1, tableSchemas
						.getString("REMARKS"));

				tabelSheet.addCell(l1);
				tabelSheet.addCell(l2);
				tabelSheet.addCell(l3);
				tabelSheet.addCell(l4);
				tabelSheet.addCell(l5);
				tabelSheet.addCell(l6);

				WritableHyperlink link = new WritableHyperlink(0, j, tableName,
						tabelSheet, 0, 0);
				sheet.addHyperlink(link);

				WritableHyperlink backLink = new WritableHyperlink(0, 0,
						tableName, sheet, 0, j);
				tabelSheet.addHyperlink(backLink);

				k = k + 1;

				// WritableHyperlink(int col, int row, java.lang.String desc,
				// WritableSheet sheet, int destcol, int
				// destrow)
				// Constructs a hyperlink to some cells within this workbook
			}

		}

		workbook.write();
		workbook.close();

	}

	// ResultSet catalogs = md.getCatalogs();
	// while (catalogs.next()) {
	// prtln(" - " + catalogs.getString(1) );
	// }
	public void readCatalogs() throws SQLException {
		System.out.println(" ==============catalog===================");
		DatabaseMetaData md = null;
		md = con.getMetaData();
		String[] names = { "TABLE" };
		ResultSet catalogs = md.getCatalogs();
		while (catalogs.next()) {
			System.out.println(catalogs.getString("TABLE_CAT"));
			// readTableColumns(metadata, table);
			// tables.addElement(table);
		}
	}

	/*
	 * get the table columns for one table
	 */
	public void readTableSchema(String tableName) throws SQLException {
		System.out
				.println(" ==============readTableSchema start===================");
		DatabaseMetaData md = null;
		md = con.getMetaData();

		tableName = "TRF_CNTR_CHG_ITEM";

		ResultSet tableSchemas = md.getColumns("", "%", tableName, "%");

		// public ResultSet getColumns(String catalog,
		// String schemaPattern,
		// String tableNamePattern,
		// String columnNamePattern)
		while (tableSchemas.next()) {
			System.out.print(tableSchemas.getString("COLUMN_NAME"));
			System.out.print(" -- ");
			System.out.print(tableSchemas.getString("DATA_TYPE"));
			System.out.print(" -- ");
			System.out.println(tableSchemas.getString("NULLABLE"));
			System.out.print(" -- ");
			System.out.print(tableSchemas.getString("COLUMN_NAME"));
			System.out.println();
		}

		System.out
				.println(" ==============readTableSchema end===================");
	}

	/*
	 * get all TMS tables
	 */
	public List readTables() throws SQLException {
		List tablesList = new ArrayList();
		System.out.println(" ==============tables===================");
		DatabaseMetaData md = null;
		md = con.getMetaData();
		String[] names = { "TABLE" };
		ResultSet tableNames = md.getTables(null, "ARPOWNER", "%", names);
		while (tableNames.next()) {
			System.out.println(tableNames.getString("TABLE_NAME"));
			tablesList.add(tableNames.getString("TABLE_NAME"));
		}

		return tablesList;
	}

	/**
	 * Sole entry point to the class and application.
	 * 
	 * @param args
	 *            Array of String arguments.
	 * @exception java.lang.InterruptedException
	 *                Thrown from the Thread class.
	 * @throws SQLException
	 * @throws IOException
	 * @throws WriteException
	 * @throws BiffException
	 * @throws RowsExceededException
	 */
	public static void main(String[] args)
			throws java.lang.InterruptedException, SQLException,
			RowsExceededException, BiffException, WriteException, IOException {

		DBSchemaForArpDev dmde = new DBSchemaForArpDev();
		dmde.writeTablesToExcel();
		System.out.println("Write out done!");

	}
}
