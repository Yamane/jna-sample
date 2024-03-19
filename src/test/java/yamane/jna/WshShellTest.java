/*
 * Copyright (c) Yamamoto Yamane
 * Released under the MIT license
 * https://opensource.org/license/mit
 */
package yamane.jna;


import static org.junit.jupiter.api.Assertions.*;

import java.io.File;

import org.junit.jupiter.api.Test;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.COM.util.ObjectFactory;

import yamane.jna.IWshShell.IWshShortcut;
import yamane.jna.WshShell.WshShortcut;

public class WshShellTest {

  @Test
  public void test1() {
    // COMライブラリの初期化
    WshShell.initialize();
    try {
      // シェルオブジェクトを取得
      WshShell shell = new WshShell();
      // ファイルの場所指定してショートカットオブジェクト生成
      WshShortcut shortcut = shell.createShortcut(shell.getSpecialFolder("Desktop") + "\\cmd1.lnk");
      // 引数
      shortcut.setArguments("/?");
      // 作業ディレクトリ
      shortcut.setWorkingDirectory(shell.getSpecialFolder("MyDocuments"));
      // リンク先（コマンドプロンプト）を設定
      shortcut.setTargetPath(System.getenv("COMSPEC"));
      // 最大化
      shortcut.setWindowStyle(3);
      // ホットキー
      shortcut.setHotkey("CTRL+SHIFT+F");
      // コメント
      shortcut.setDescription("ショートカットサンプル1");
      // アイコン
      shortcut.setIconLocation("C:\\Windows\\System32\\calc.exe, 0");
      // 保存
      shortcut.save();
      
      // 一応テスト
      File file = new File(shell.getSpecialFolder("Desktop") + "\\cmd1.lnk");
      assertTrue(file.exists());
      file.delete();
      
    } finally {
      // COMライブラリの初期化解除
      WshShell.unInitialize();
    }
    
  }

  @Test
  public void test2() {
    // COMライブラリの初期化
    Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
    try {
      // ファクトリを準備
      ObjectFactory factory = new ObjectFactory();
      // シェルオブジェクトを取得
      IWshShell shell = factory.createObject(IWshShell.class);
      // ファイルの場所指定してショートカットオブジェクト生成
      String linkpath = shell.getSpecialFolder("Desktop") + "\\cmd2.lnk";
      IDispatch dispatch = shell.createShortcut(linkpath);
      IWshShortcut shortcut = factory.createProxy(IWshShortcut.class, dispatch);
      // 引数
      shortcut.setArguments("/?");
      // 作業ディレクトリ
      shortcut.setWorkingDirectory(shell.getSpecialFolder("MyDocuments"));
      // リンク先（コマンドプロンプト）を設定
      shortcut.setTargetPath(System.getenv("COMSPEC"));
      // 最大化
      shortcut.setWindowStyle(3);
      // ホットキー
      shortcut.setHotkey("CTRL+SHIFT+F");
      // コメント
      shortcut.setDescription("ショートカットサンプル2");
      // アイコン
      shortcut.setIconLocation("C:\\Windows\\System32\\notepad.exe, 0");
      // 保存
      shortcut.save();
      
      // 一応テスト
      File file = new File(shell.getSpecialFolder("Desktop") + "\\cmd2.lnk");
      assertTrue(file.exists());
      file.delete();
      
    } finally {
      // COMライブラリの初期化解除
      Ole32.INSTANCE.CoUninitialize();
    }
  }
}
