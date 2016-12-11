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
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.Timer;
import javax.swing.JFrame;
import javax.swing.JPanel;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.time.Day;
import org.jfree.data.time.FixedMillisecond;
import org.jfree.data.time.Millisecond;
import org.jfree.data.time.RegularTimePeriod;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.data.xy.XYDataset;
import org.jfree.ui.ApplicationFrame;
import org.jfree.ui.RefineryUtilities;


public class serialRecorder extends ApplicationFrame implements SerialPortEventListener, ActionListener {

	// excel sheet
	static HSSFSheet sheet;
	static boolean print = true
			;
	
	/** The time series data. */
    private TimeSeries series;

    /** The most recent value added. */
    private double lastValue = 100.0;
   
    /** Timer to refresh graph after every 1/4th of a second */
    private Timer timer = new Timer(250, this);

	public serialRecorder(final String title) {
		super(title);
        this.series = new TimeSeries("Acceleration X", Millisecond.class);
       
        final TimeSeriesCollection dataset = new TimeSeriesCollection(this.series);
        final JFreeChart chart = createChart(dataset);
       
        timer.setInitialDelay(1000);
       
        //Sets background color of chart
        chart.setBackgroundPaint(Color.LIGHT_GRAY);
       
        //Created JPanel to show graph on screen
        final JPanel content = new JPanel(new BorderLayout());
       
        //Created Chartpanel for chart area
        final ChartPanel chartPanel = new ChartPanel(chart);
       
        //Added chartpanel to main panel
        content.add(chartPanel);
        
        //Sets the size of whole window (JPanel)
        chartPanel.setPreferredSize(new java.awt.Dimension(800, 500));
       
        //Puts the whole content on a Frame
        setContentPane(content);
      
        
        timer.start();
	}
	
	 /**
     * Creates a sample chart.
     *
     * @param dataset  the dataset.
     *
     * @return A sample chart.
     */
    private JFreeChart createChart(final XYDataset dataset) {
        final JFreeChart result = ChartFactory.createTimeSeriesChart(
            "Acceleration vs. Time",
            "Time",
            "g's",
            dataset,
            true,
            true,
            false
        );
       
        final XYPlot plot = result.getXYPlot();
       
        plot.setBackgroundPaint(new Color(0xffffe0));
        plot.setDomainGridlinesVisible(true);
        plot.setDomainGridlinePaint(Color.lightGray);
        plot.setRangeGridlinesVisible(true);
        plot.setRangeGridlinePaint(Color.lightGray);
                
        ValueAxis xaxis = plot.getDomainAxis();
        xaxis.setAutoRange(true);
       
        //Domain axis would show data of 60 seconds for a time
        xaxis.setFixedAutoRange(5000.0);  // 60 seconds
        xaxis.setVerticalTickLabels(true);
       
        ValueAxis yaxis = plot.getRangeAxis();
        yaxis.setRange(0.0, 16.0);
       
        return result;
    }
    /**
     * Generates an random entry for a particular call made by time for every 1/4th of a second.
     *
     * @param e  the action event.
     */
    /*
    public void actionPerformed(final ActionEvent e) {
       
        final double factor = 0.9 + 0.2*Math.random();
        this.lastValue = this.lastValue * factor;
       
        final Millisecond now = new Millisecond();
        this.series.add(new Millisecond(), this.lastValue);
       
        System.out.println("Current Time in Milliseconds = " + now.toString()+", Current Value : "+this.lastValue);
    }*/

	
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
		final serialRecorder demo = new serialRecorder("");
        demo.pack();
        RefineryUtilities.centerFrameOnScreen(demo);
        demo.setVisible(true);
        
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
		demo.initialize();
		System.out.println("\nStarted Recording...(Press ENTER to stop recording)");
		

		in.nextLine();
		print = false;
		
		System.out.println("Max X Acceleration: " + maxX + " g's");

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
		System.exit(0);
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
	long time = 0;
	static // "x10, y10, z10
	double maxX = 0;
	int packets = 0;
	
	double lastVal = 0;
	
	@Override
	public void serialEvent(SerialPortEvent oEvent) {
		if (oEvent.getEventType() == SerialPortEvent.DATA_AVAILABLE) {
			try {
				 while (input.ready ())
	              {
	              
				String inputLine = input.readLine();

				// System.out.println("swag");

				if (print && (inputLine.startsWith("x") || inputLine.startsWith("y") || inputLine.startsWith("z")) && (System.currentTimeMillis() - time > 5)) {
					//System.out.println("Received " + packets + " samples - total: " + r + "... (Press ENTER to stop recording)");
					time = System.currentTimeMillis();
					this.series.add(new Millisecond(), lastVal);
					packets = 0;
				}
				
				if (inputLine.startsWith("x")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row = sheet.createRow(r);
					row.createCell(0).setCellValue(val);
					
					lastVal = val;
					
					if (val > maxX) {
						maxX = val;
					}
					packets++;
					
					r++;
				} else if (inputLine.startsWith("y")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row.createCell(1).setCellValue(val);
				} else if (inputLine.startsWith("z")) {
					double val = Double.parseDouble(inputLine.substring(1));
					row.createCell(2).setCellValue(val);

					/*if (print) {
						System.out.printf("x: %-15.2f y: %-15.2f z: %-15.2f \t (Press ENTER to stop recording)\n",
								row.getCell(0).getNumericCellValue(), row.getCell(1).getNumericCellValue(),
								row.getCell(2).getNumericCellValue());
					}*/
				}
	              }
			} catch (Exception e) {
				System.err.println(e.toString());
			}
		}
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		
	}

}
