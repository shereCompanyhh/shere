package testSample;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.security.InvalidAlgorithmParameterException;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;

import javax.crypto.BadPaddingException;
import javax.crypto.Cipher;
import javax.crypto.IllegalBlockSizeException;
import javax.crypto.NoSuchPaddingException;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;


import org.apache.commons.codec.binary.Base64;


public class Encdata {

	 private static String ENCRYPT_KEY = ""; 
	 private static String ENCRYPT_IV = "";

	/**
	 * メインメソッド
	 *
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		System.out.println("暗号化復号化ツールを実行しいます。");
		if(args.length == 5) {
			System.out.println("引数の数が不正です。６つ指定してください");
			System.exit(0);
		}
		
		if(args[0].trim().length() == 0) {
			System.out.println("第１引数の読み込みファイルが指定されていません。");
			System.exit(0);
		}
		if(args[1].trim().length() == 0) {
			System.out.println("第２引数の出力先が指定されていません。");
			System.exit(0);
		}
		if(args[2].trim().length() == 0) {
			System.out.println("第３引数の出力ファイル名が指定されていません。");
			System.exit(0);
		}
		
		if(args[3].trim().length() == 0) {
			System.out.println("第3引数のモードが指定されていません。");
			System.exit(0);
		}
		
		if(args[4].trim().length() == 0) {
			System.out.println("第４引数の暗号・複合キーが指定されていません。");
			System.exit(0);
		}
		
		if(args[5].trim().length() == 0) {
			System.out.println("第５引数の暗号・複合IVが指定されていません。");
			System.exit(0);
		}
		
		
		String inputfilepath = args[0];
		String outputfilepath = args[1];
		String outputfilename = args[2];
		String mode = args[3];
		ENCRYPT_KEY = args[4];
		ENCRYPT_IV = args[5];

		
		if(mode.equals("0")) {
			// 復号化メソッド呼出
			dec_base64(decrypt(read(inputfilepath)),outputfilepath + "\\" +outputfilename);
		}else {
			// 暗号化メソッド呼出
			write(encrypt(inputfile_base64(inputfilepath).toString()),outputfilepath + "\\" +outputfilename);
		}
	}

	private static String read(String readfilepath) throws IOException {
	
		File file = new File(readfilepath);
		String rs = "";
	      if (checkBeforeReadfile(file)){
	        BufferedReader br = new BufferedReader(new FileReader(file));

	        String str = "";
	        while((str = br.readLine()) != null){
	          System.out.println(str);
	          rs = rs + str;
	        }

	        br.close();
	      }else{
	        System.out.println("ファイルが見つからないか開けません");
	        return rs;
	      }
        return rs;
	}
	
	  private static boolean checkBeforeReadfile(File file){
		    if (file.exists()){
		      if (file.isFile() && file.canRead()){
		        return true;
		      }
		    }

		    return false;
		  }

	private static void write(String str,String filepath) throws IOException {
        // FileWriterクラスのオブジェクトを生成する
        FileWriter file = new FileWriter(filepath);
        // PrintWriterクラスのオブジェクトを生成する
        PrintWriter pw = new PrintWriter(new BufferedWriter(file));
        
        //ファイルに書き込む
        pw.print(str);
        pw.flush();

        
        //ファイルを閉じる
        pw.close();
		
		
	}
	private static StringBuffer inputfile_base64(String inputFilePath) throws IOException {
		File file = new File(inputFilePath);
		int fileLen = (int)file.length();
		byte[] data = new byte[fileLen];
		FileInputStream fis = new FileInputStream(file);
		fis.read(data);
		StringBuffer rs = new StringBuffer();
		rs.setLength(0);
		rs.append(Base64.encodeBase64String(data));
		
		System.out.println("bse64データ");
		System.out.println(rs.toString());
		
		
		
		return rs;
	}

	private static void dec_base64(String str,String outputfilePath) throws IOException {
		byte[] data2 = Base64.decodeBase64(str);
		FileOutputStream fos = new FileOutputStream(outputfilePath);
		fos.write(data2);
		fos.flush();
		fos.close();
	}
	
	/**
	 * 暗号化メソッド
	 *
	 * @param text 暗号化する文字列
	 * @return 暗号化文字列
	 */
	public static String encrypt(String text) {
		// 変数初期化
		String strResult = null;

		try {
			// 文字列をバイト配列へ変換
			byte[] byteText = text.getBytes("UTF-8");

			// 暗号化キーと初期化ベクトルをバイト配列へ変換
			byte[] byteKey = ENCRYPT_KEY.getBytes("UTF-8");
			byte[] byteIv = ENCRYPT_IV.getBytes("UTF-8");

			// 暗号化キーと初期化ベクトルのオブジェクト生成
			SecretKeySpec key = new SecretKeySpec(byteKey, "AES");
			IvParameterSpec iv = new IvParameterSpec(byteIv);

			// Cipherオブジェクト生成
			Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");

			// Cipherオブジェクトの初期化
			cipher.init(Cipher.ENCRYPT_MODE, key, iv);

			// 暗号化の結果格納
			byte[] byteResult = cipher.doFinal(byteText);

			// Base64へエンコード
			strResult = Base64.encodeBase64String(byteResult);

		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (NoSuchPaddingException e) {
			e.printStackTrace();
		} catch (InvalidKeyException e) {
			e.printStackTrace();
		} catch (IllegalBlockSizeException e) {
			e.printStackTrace();
		} catch (BadPaddingException e) {
			e.printStackTrace();
		} catch (InvalidAlgorithmParameterException e) {
			e.printStackTrace();
		}

		// 暗号化文字列を返却
		return strResult;
	}

	/**
	 * 復号化メソッド
	 *
	 * @param text 復号化する文字列
	 * @return 復号化文字列
	 */
	public static String decrypt(String text) {
		// 変数初期化
		String strResult = null;

		try {
			// Base64をデコード
			byte[] byteText = Base64.decodeBase64(text);

			// 暗号化キーと初期化ベクトルをバイト配列へ変換
			byte[] byteKey = ENCRYPT_KEY.getBytes("UTF-8");
			byte[] byteIv = ENCRYPT_IV.getBytes("UTF-8");

			// 復号化キーと初期化ベクトルのオブジェクト生成
			SecretKeySpec key = new SecretKeySpec(byteKey, "AES");
			IvParameterSpec iv = new IvParameterSpec(byteIv);

			// Cipherオブジェクト生成
			Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");

			// Cipherオブジェクトの初期化
			cipher.init(Cipher.DECRYPT_MODE, key, iv);

			// 復号化の結果格納
			byte[] byteResult = cipher.doFinal(byteText);

			// バイト配列を文字列へ変換
			strResult = new String(byteResult, "UTF-8");

		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (NoSuchPaddingException e) {
			e.printStackTrace();
		} catch (InvalidKeyException e) {
			e.printStackTrace();
		} catch (IllegalBlockSizeException e) {
			e.printStackTrace();
		} catch (BadPaddingException e) {
			e.printStackTrace();
		} catch (InvalidAlgorithmParameterException e) {
			e.printStackTrace();
		}

		// 復号化文字列を返却
		return strResult;
	}
}