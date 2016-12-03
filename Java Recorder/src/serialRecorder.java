import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Paths;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import gnu.io.CommPortIdentifier;
import gnu.io.SerialPort;
import gnu.io.SerialPortEvent;
import gnu.io.SerialPortEventListener;

import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Scanner;

public class serialRecorder implements SerialPortEventListener {

	// excel sheet
	static HSSFSheet sheet;
	static boolean print = true;

	// Serial shit
	SerialPort serialPort;
	/** The port we're normally going to use. */
	private static final String PORT_NAMES[] = { "/dev/tty.usbserial-A9007UX1", "/dev/cu.usbmodem1451", // Mac
																										// //
																										// X
			"/dev/ttyACM0", // Raspberry Pi
			"/dev/ttyUSB0", // Linux
			"COM3", // Windows
	};
	/**
	 * A BufferedReader which will be fed by a InputStreamReader converting the
	 * bytes into characters making the displayed results codepage independent
	 */
	private BufferedReader input;
	/** Milliseconds to block while waiting for port open */
	private static final int TIME_OUT = 2000;
	/** Default bits per second for COM port. */
	private static final int DATA_RATE = 115200;

	public void initialize() {

		CommPortIdentifier portId = null;

		ArrayList<CommPortIdentifier> portArray = new ArrayList<>();
		
		System.out.println("COMM Ports");
		// First, Find an instance of serial port as set in PORT_NAMES.
		
		
		
		while (portId == null) {
			
			
			Enumeration<?> portEnum = CommPortIdentifier.getPortIdentifiers();
			portArray.clear();
			while (portEnum.hasMoreElements()) {
				CommPortIdentifier currPortId = (CommPortIdentifier) portEnum.nextElement();
				portArray.add(currPortId);
			}
			
			System.out.println();
			int c = 0;
			for (CommPortIdentifier cpi : portArray) {
				System.out.println("(" + c + ") - "  + cpi.getName());
				c++;
			}
			System.out.print("select a port #, or (r)efresh, (q)uit: ");
			
			String r = in.nextLine();
			if (r.contains("q")) {
				System.exit(0);
			} else if (!r.contains("r")) {
				int i = Integer.parseInt(r);
				portId = portArray.get(i);
			}
			
			
		}
		
		if (portId == null) {
			System.out.println("\nCould not find COM port.");
			System.exit(0);
			return;
		}

		try {
			// open serial port, and use class name for the appName.
			serialPort = (SerialPort) portId.open(this.getClass().getName(), TIME_OUT);

			// set port parameters
			serialPort.setSerialPortParams(DATA_RATE, SerialPort.DATABITS_8, SerialPort.STOPBITS_1,
					SerialPort.PARITY_NONE);

			// open the streams
			input = new BufferedReader(new InputStreamReader(serialPort.getInputStream()));
			serialPort.getOutputStream();

			// add event listeners
			serialPort.addEventListener(this);
			serialPort.notifyOnDataAvailable(true);
		} catch (Exception e) {
			System.err.println(e.toString());
		}
	}

	static Scanner in = new Scanner(System.in);

	public static void main(String[] args) throws IOException, URISyntaxException {
		HSSFWorkbook workbook = null;
		try {
			workbook = new HSSFWorkbook();
			sheet = workbook.createSheet("FirstSheet");

			HSSFRow rowhead = sheet.createRow((short) 0);
			rowhead.createCell(0).setCellValue("Accel X");
			rowhead.createCell(1).setCellValue("Accel Y");
			rowhead.createCell(2).setCellValue("Accel Z");

		} catch (Exception ex) {
			System.out.println(ex);
		}

		serialRecorder main = new serialRecorder();
		main.initialize();
		Thread t = new Thread() {
			public void run() {
				// the following line will keep this app alive for 1000 seconds,
				// waiting for events to occur and responding to them (printing
				// incoming messages to console).
				try {
					while (print) {
						Thread.sleep(1000);
					}
				} catch (InterruptedException ie) {
				}
			}
		};
		t.start();
		System.out.println("\nStarted Recording...(Press ENTER to stop recording)");
		

		in.nextLine();
		print = false;

		System.out.print("(s)ave or (q)uit? ");
		String s = in.nextLine();
		
		if (s.contains("s")) {
			System.out.print("File name: ");

			String name = in.nextLine();

			
			String path = Paths.get(serialRecorder.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent().toString();
			
			String folder = "ardu excel";
			
			File f = new File(path + "/" + folder);
			
			if (!f.exists()) {
				f.mkdirs();
			}
			
			FileOutputStream fileOut = new FileOutputStream(path + "/" + folder + "/" + name + ".xls");
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("Saving excel file with name: " + name + ".xls");
		} else {
			System.exit(0);
		}

		main.close();
	}

	public synchronized void close() {
		if (serialPort != null) {
			serialPort.removeEventListener();
			serialPort.close();
		}
	}

	short r = 1;
	int c = 0;
	HSSFRow row;

	// "x10, y10, z10
	@Override
	public void serialEvent(SerialPortEvent oEvent) {
		if (oEvent.getEventType() == SerialPortEvent.DATA_AVAILABLE) {
			try {
				String inputLine = input.readLine();

				// System.out.println("swag");

				if (inputLine.startsWith("x")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row = sheet.createRow(r);
					row.createCell(0).setCellValue(val);
					r++;
				} else if (inputLine.startsWith("y")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row.createCell(1).setCellValue(val);
				} else if (inputLine.startsWith("z")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row.createCell(2).setCellValue(val);

					if (print) {
						System.out.printf("x: %-15.2f y: %-15.2f z: %-15.2f \t (Press ENTER to stop recording)\n",
								row.getCell(0).getNumericCellValue(), row.getCell(1).getNumericCellValue(),
								row.getCell(2).getNumericCellValue());
					}
				}

			} catch (Exception e) {
				System.err.println(e.toString());
			}
		}
	}

}
