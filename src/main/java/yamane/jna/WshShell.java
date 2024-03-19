/*
 * Copyright (c) Yamamoto Yamane
 * Released under the MIT license
 * https://opensource.org/license/mit
 */

package yamane.jna;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * WshShellクラス
 */
public class WshShell extends COMLateBindingObject {

  /**
   * コンストラクタ.
   */
  public WshShell() {
    super("WScript.Shell", false);
  }

  /**
   * 現在アクティブになっているディレクトリを取得します。
   * @return 現在アクティブになっているディレクトリのパス
   */
  public String getCurrentDirectory() {
    return getStringProperty("CurrentDirectory");
  }

  /**
   * SpecialFolders プロパティ(コレクション)から特定の要素を取得します。
   * @param name 要素名
   * @return フォルダーのパス
   */
  public String getSpecialFolder(String name) {
    VARIANT val = invoke("SpecialFolders", new VARIANT(name));
    return val.getValue().toString();
  }

  /**
   * ショートカットオブジェクトを生成します。
   * @param path 生成するショートカットのパス
   * @return ショートカットオブジェクト
   */
  public WshShortcut createShortcut(String path) {
    VARIANT val = invoke("CreateShortcut", new VARIANT(path));
    return new WshShortcut((IDispatch) val.getValue());
  }

  /**
   * USER環境変数を取得します。
   * @param name 変数名
   * @return 変数値
   */
  public String getUserEnvironment(String name) {
    return getEnvironment("USER").item(name);
  }

  /**
   * SYSTEM環境変数を取得します。
   * @param name 変数名
   * @return 変数値
   */
  public String getSystemEnvironment(String name) {
    return getEnvironment("SYSTEM").item(name);
  }

  /**
   * WshEnvironment オブジェクト (環境変数のコレクション) を返します。
   * @param name 環境変数の種類
   * @return 環境変数のコレクション
   */
  private WshEnvironment getEnvironment(String name) {
    VARIANT val = invoke("Environment", new VARIANT(name));
    return new WshEnvironment((IDispatch) val.getValue());
  }

  /**
   * ポップアップ メッセージ ボックスにテキストを表示します。
   * @param message ポップアップ ウィンドウに表示するテキスト
   * @param secondsToWait ポップアップ ウィンドウを閉じるまで待機する秒数
   * @param title ポップアップ ウィンドウのタイトルに表示するテキスト
   * @param type ボタンとアイコンの種類を示す数値
   * @return メッセージ ボックス終了時にクリックするボタンの番号を示す整数値
   */
  public Number popup(String message, Integer secondsToWait, String title, Integer type) {
    VARIANT[] valiants = new VARIANT[] {
        new VARIANT(message),
        new VARIANT(secondsToWait),
        new VARIANT(title),
        new VARIANT(type),
    };
    VARIANT val = invoke("Popup", valiants);
    return (Number) val.getValue();
  }

  /**
   * 新しいプロセス内でプログラムを実行します。
   * @param command 実行するコマンドライン
   */
  public void run(String command) {
    invoke("Run", new VARIANT(command));
  }

  /**
   * 新しいプロセス内でプログラムを実行します。
   * @param command 実行するコマンドライン
   * @param style プログラムのウィンドウの外観を示す整数値
   * @param isWait プログラムの実行が終了するまでスクリプトを待機させるかどうか
   */
  public void run(String command, Integer style, Boolean isWait) {
    invoke("Run", new VARIANT(command), new VARIANT(style), new VARIANT(isWait));
  }

  /**
   * WshShortcutクラス
   */
  public class WshShortcut extends COMLateBindingObject {

    private WshShortcut(IDispatch iDispatch) {
      super(iDispatch);
    }

    /**
     * ショートカットの実行可能ファイルへのパスを設定します。
     * @param path ショートカットの実行可能ファイルへのパス
     */
    public void setTargetPath(String path) {
      setProperty("TargetPath", path);
    }

    /**
     * ショートカット実行時の引数を設定します。
     * @param args ショートカット実行時の引数
     */
    public void setArguments(String args) {
      setProperty("Arguments", args);
    }

    /**
     * ショートカットの作業ディレクトリを設定します。
     * @param dir 作業ディレクトリのパス
     */
    public void setWorkingDirectory(String dir) {
      setProperty("WorkingDirectory", dir);
    }

    /**
     * アイコンをショートカットに割り当てます。
     * 絶対パスと、アイコンに関連付けられているインデックスを含める必要があります。
     * @param path アイコンの格納場所を示すパス+インデックス
     */
    public void setIconLocation(String path) {
      setProperty("IconLocation", path);
    }

    /**
     * ショートカットに割り当てるキーの組み合わせを設定します。
     * @param keys キーの組み合わせを表す文字列
     */
    public void setHotkey(String keys) {
      setProperty("Hotkey", keys);
    }

    /**
     * ウィンドウ スタイルをショートカットに割り当てます。
     * <ul>
     * <li>1. ウィンドウをアクティブにして表示</li>
     * <li>3. ウィンドウを最大化して表示</li>
     * <li>7. ウィンドウを最小化して表示</li>
     * </ul>
     * @param style キーの組み合わせを表す文字列
     */
    public void setWindowStyle(Integer style) {
      setProperty("WindowStyle", style);
    }

    /**
     * ショートカットの説明文（コメント）を設定します。
     * @param text 説明文（コメント）
     */
    public void setDescription(String text) {
      setProperty("Description", text);
    }

    /**
     * ショートカットオブジェクトを保存します。
     */
    public void save() {
      invoke("Save");
    }
  }

  /**
   * WshEnvironmentクラス
   */
  private class WshEnvironment extends COMLateBindingObject {

    private WshEnvironment(IDispatch iDispatch) {
      super(iDispatch);
    }

    /**
     * コレクションから、指定されたアイテムを返します。
     * @param name アイテム名
     * @return アイテム値
     */
    private String item(String name) {
      return invoke("Item", new VARIANT(name)).getValue().toString();
    }
  }

  /**
   * COMライブラリの初期化
   */
  public static void initialize() {
    Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
  }

  /**
   * COMライブラリの初期化解除
   */
  public static void unInitialize() {
    Ole32.INSTANCE.CoUninitialize();
  }
}
