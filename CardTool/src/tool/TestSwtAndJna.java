package tool;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import net.sf.json.JSONObject;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.custom.StackLayout;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;

public class TestSwtAndJna {
    static OleClientSite clientSite;
    
    static String CSN = null;
    
    static String data = "";

    public static void main(String[] args) {
        Display display = new Display();

        Shell shell = new Shell(display);
        shell.setText("测试调用ActiveX来处理卡片");
        shell.setLayout(new FillLayout(SWT.VERTICAL));

        try {
            OleFrame frame = new OleFrame(shell, SWT.NONE);
           StackLayout layout =  new StackLayout();
            frame.setLayout(layout);
            text  = new Text(frame, SWT.MULTI|SWT.WRAP|SWT.V_SCROLL);
            layout.topControl = text;
            clientSite = new OleClientSite(frame, SWT.NONE,
                    "{711FF4B0-884F-42E1-978B-CD19AEB3BC3E}");
            clientSite.setVisible(false);
            clientSite.setSize(0,0);
            clientSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
            addFileMenu(frame);

        } catch (SWTError e) {
        	showInfoMessageBox("打开Active控件失败", shell);
            display.dispose();
            return;
        }
        catch(Exception e){
        	showInfoMessageBox("打开Active控件失败", shell);
            display.dispose();
            return;
        }
        shell.setSize(400,300);
        shell.open();
        while (!shell.isDisposed()) {
            if (!display.readAndDispatch()) {
                display.sleep();
            }
        }

        display.dispose();
    }
    
    private static Text text;

    static void addFileMenu(final OleFrame frame) {
        final Shell shell = frame.getShell();
        Menu menuBar = shell.getMenuBar();
        if (menuBar == null) {
            menuBar = new Menu(shell, SWT.BAR);
            shell.setMenuBar(menuBar);
        }
        MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
        fileMenu.setText("&Action");
        Menu menuFile = new Menu(fileMenu);
        fileMenu.setMenu(menuFile);
        frame.setFileMenus(new MenuItem[] { fileMenu });

        MenuItem menuHandle = new MenuItem(menuFile, SWT.CASCADE);
        menuHandle.setText("查看ActiveX窗口的句柄");
        menuHandle.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
                MessageBox messageBox = new MessageBox(shell,
                        SWT.ICON_INFORMATION | SWT.OK);
                messageBox.setText("Info");

                // useful with SpyXX.exe
                messageBox.setMessage("handle = ["
                        + Long.toHexString(clientSite.handle).toUpperCase()
                        + "]");
                messageBox.open();
            }
        });
        
        MenuItem openComMenuTextData = new MenuItem(menuFile, SWT.CASCADE);
        openComMenuTextData.setText("打开串口3");
        openComMenuTextData.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	try{
	                OleAutomation ole = new OleAutomation(clientSite);
	                
	                int[] rgdispid = ole.getIDsOfNames(new String[] { "OpenCOM"});
	                if (rgdispid == null || rgdispid.length == 0) {
	                	appendContent("找不到此方法");
	                    return;  
	                }
	                //构造参数列表  
	                Variant[] variants = new Variant[1];  
	                variants[0] = new Variant(3);  
	          
	                Variant pVarResult = ole.invoke(rgdispid[0], variants);  
	                 
	                //释放参数对象  
	                for (int i = 0; i < variants.length; i++) {  
	                  variants[i].dispose();  
	                }  
	                 
	                //获取调用结果  
	                if (pVarResult != null) {  
	                  Object value = getValue(pVarResult);  
	                  pVarResult.dispose(); 
	                  appendContent("打开串口，返回："+ value);
	                } 
            	}catch(Exception ex){
            		ex.printStackTrace();
            	}
            }
        });
        
        MenuItem closeComMenuTextData = new MenuItem(menuFile, SWT.CASCADE);
        closeComMenuTextData.setText("关闭串口3");
        closeComMenuTextData.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	try{
	                OleAutomation ole = new OleAutomation(clientSite);
	                
	                int[] rgdispid = ole.getIDsOfNames(new String[] { "CloseCOM"});
	                if (rgdispid == null || rgdispid.length == 0) {
	                	appendContent("找不到此方法");
	                    return;  
	                }
	                //构造参数列表  
	                Variant[] variants = new Variant[1];  
	                variants[0] = new Variant(3);  
	          
	                Variant pVarResult = ole.invoke(rgdispid[0], variants);  
	                 
	                //释放参数对象  
	                for (int i = 0; i < variants.length; i++) {  
	                  variants[i].dispose();  
	                }  
	                 
	                //获取调用结果  
	                if (pVarResult != null) {  
	                  Object value = getValue(pVarResult);  
	                  pVarResult.dispose(); 
	                  appendContent("关闭串口，返回："+ value);
	                } 
            	}catch(Exception ex){
            		ex.printStackTrace();
            	}
            }
        });

        MenuItem getCardInfoMenuTextData = new MenuItem(menuFile, SWT.CASCADE);
        getCardInfoMenuTextData.setText("读取卡信息");
        getCardInfoMenuTextData.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	try{
	                OleAutomation ole = new OleAutomation(clientSite);
	                int[] rgdispid = ole.getIDsOfNames(new String[] { "GetCardInfo"});
	                if (rgdispid == null || rgdispid.length == 0) {
	                	appendContent("找不到此方法");
	                    return;  
	                }
	                Variant pVarResult = ole.invoke(rgdispid[0]);    
	                //获取调用结果  
	                if (pVarResult != null) {
	                  Object value = getValue(pVarResult);  
	                  pVarResult.dispose(); 
	                  appendContent("读卡成功，返回："+ value);
	                  String result = (String)value;
	                  if(value!=null){
	                	  
	                	  CSN = result.substring(17,25);
	                	  appendContent(CSN);
	                	  try{
	                		  JSONObject json = JSONObject.fromObject(result);
		                	  for (Object o : json.keySet()) {
								appendContent("key: " + o +"  value: "+json.get(o));
		                	  }
	                	  }catch (Exception ex) {
							ex.printStackTrace();
						  }
	                	  
	                	  
	                  }
	                } 
            	}catch(Exception ex){
            		ex.printStackTrace();
            	}
            }
        });

        MenuItem setCardToException = new MenuItem(menuFile, SWT.CASCADE);
        setCardToException.setText("设置为异常卡");
        setCardToException.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	if(CSN == null){
            		appendContent("没有CSN号，不可设置为异常卡");
            		return;
            	}
            	OleAutomation ole = new OleAutomation(clientSite);
                int[] rgdispid = ole.getIDsOfNames(new String[] { "AbnormalDeal"});
                if (rgdispid == null || rgdispid.length == 0) {
                	appendContent("找不到方法 : AbnormalDeal");
                    return;
                }
                //构造参数列表 
                Variant[] variants = new Variant[3];  
                variants[0] = new Variant(getChipType());
                variants[1] = new Variant(CSN);
                variants[2] = new Variant("1308200830" + "1308200930" + "00012345" + "A5");
                Variant pVarResult = ole.invoke(rgdispid[0], variants);  
                 
                //释放参数对象  
                for (int i = 0; i < variants.length; i++) {  
                  variants[i].dispose();  
                }     
                //获取调用结果  
                if (pVarResult != null) {
                  Object value = getValue(pVarResult);  
                  pVarResult.dispose(); 
                  appendContent("改成异常卡成功，返回："+ value);
                } 
            }
        });

        MenuItem setCardToOK = new MenuItem(menuFile, SWT.CASCADE);
        setCardToOK.setText("设置为正常卡");
        setCardToOK.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	if(CSN == null){
            		appendContent("没有CSN号，不可设置为正常卡");
            		return;
            	}
            	OleAutomation ole = new OleAutomation(clientSite);
                int[] rgdispid = ole.getIDsOfNames(new String[] { "AbnormalDeal"});
                if (rgdispid == null || rgdispid.length == 0) {
                	appendContent("找不到方法 : AbnormalDeal");
                    return;
                }
                //构造参数列表 
                Variant[] variants = new Variant[3];  
                variants[0] = new Variant(getChipType());
                variants[1] = new Variant(CSN);
                variants[2] = new Variant("1308200830" + "1308200930" + "12345678" + "96");
                Variant pVarResult = ole.invoke(rgdispid[0], variants);  
                 
                //释放参数对象  
                for (int i = 0; i < variants.length; i++) {  
                  variants[i].dispose();  
                }     
                //获取调用结果  
                if (pVarResult != null) {
                  Object value = getValue(pVarResult);  
                  pVarResult.dispose(); 
                  appendContent("改成异常卡成功，返回："+ value);
                } 
            }
        });
        
        MenuItem consumptionMenuItem = new MenuItem(menuFile, SWT.CASCADE);
        consumptionMenuItem.setText("消费0.01元");
        consumptionMenuItem.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent e) {
            	if(CSN == null){
            		appendContent("没有CSN号，不可消费");
            		return;
            	}
            	OleAutomation ole = new OleAutomation(clientSite);
                int[] rgdispid = ole.getIDsOfNames(new String[] { "Consumption"});
                if (rgdispid == null || rgdispid.length == 0) {
                	appendContent("找不到方法 : Consumption");
                    return;
                }
                //构造参数列表 
                Variant[] variants = new Variant[4];  
                variants[0] = new Variant(getChipType());
                variants[1] = new Variant(CSN);
                variants[2] = new Variant(123);
                SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
                variants[3] = new Variant(format.format(new Date()));
                Variant pVarResult = ole.invoke(rgdispid[0], variants);  
                 
                //释放参数对象  
                for (int i = 0; i < variants.length; i++) {  
                  variants[i].dispose();  
                }     
                //获取调用结果  
                if (pVarResult != null) {
                  Object value = getValue(pVarResult);  
                  pVarResult.dispose(); 
                  appendContent("消费一分钱成功，返回："+ value);
                } 
            }
        });
    }
    
    private static Object getValue(Variant variant) {  
        short type = variant.getType();  
      
        switch (type) {  
        case COM.VT_BOOL:  
            return variant.getBoolean();  
        case COM.VT_I2:  
            return variant.getShort();  
        case COM.VT_I4:  
            return variant.getInt();  
        case COM.VT_R4:  
            return variant.getFloat();  
        case COM.VT_BSTR:  
            return variant.getString();  
        case COM.VT_DISPATCH:  
            return variant.getAutomation();  
        case COM.VT_UNKNOWN:  
            return variant;  
        case COM.VT_EMPTY:  
            return null;  
        }  
        if ((type & COM.VT_BYREF) != 0) {  
            return variant.getByRef();  
        }  
      
        return null;  
      }
    private static void appendContent(String content){
    	text.append(content+ "\n");
    }
    
    public static void showInfoMessageBox(String msg, Shell shell){
    	MessageBox messageBox = new MessageBox(shell, SWT.OK|SWT.ICON_INFORMATION);
    	messageBox.setMessage(msg);
    	messageBox.setText("提示");
    	messageBox.open();
    	
    }
    
    
	public static Integer getChipType() {
		Map<String, String> cardInfo = GetCardInfo();
		String xplxz = cardInfo.get("xplxz");
		return Integer.parseInt(xplxz);
	}
	
	
	public static Map<String,String> GetCardInfo(){
		Map<String,String> map = new HashMap<String, String>(); 
		try{
            OleAutomation ole = new OleAutomation(clientSite);
            int[] rgdispid = ole.getIDsOfNames(new String[] { "GetCardInfo"});
            if (rgdispid == null || rgdispid.length == 0) {
            	appendContent("严重错误，无法找到读卡信息的相应函数，请检查OCX控件是否正常注册！");
                return null;  
            }
            Variant pVarResult = ole.invoke(rgdispid[0]);    
            //获取调用结果  
            if (pVarResult != null) {
            	String result = pVarResult.getString();
				pVarResult.dispose(); 
				if(result!=null){
					try {
						JSONObject json = JSONObject.fromObject(result);
						map.putAll(json);
					} catch (Exception ex) {
						ex.printStackTrace();
						map.put("Err", "结果集类型转换异常！");
					}
              }else{
            	  map.put("Err", "没读到卡信息！");
              }
            } 
    	}catch(Exception ex){
    		ex.printStackTrace();
    		map.put("Err", "读卡时出现异常！");
    	}
	
		
		return map;
	}
}

