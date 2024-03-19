/*
 * Copyright (c) Yamamoto Yamane
 * Released under the MIT license
 * https://opensource.org/license/mit
 */

package yamane.jna;

import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.COM.util.annotation.ComMethod;
import com.sun.jna.platform.win32.COM.util.annotation.ComObject;
import com.sun.jna.platform.win32.COM.util.annotation.ComProperty;

/**
 * WshShellインターフェイス.
 */
@ComObject(progId = "WScript.Shell")
public interface IWshShell {

	/**
	 * SpecialFolders プロパティ(コレクション)から特定の要素を取得します.
	 * @param name 要素名
	 * @return フォルダーのパス
	 */
	@ComProperty(name = "SpecialFolders")
	String getSpecialFolder(String name);

	/**
	 * ショートカットオブジェクトを生成します.
	 * @param path 生成するショートカットのパス
	 * @return WshShortcutインターフェイスを持つオブジェクト
	 */
	@ComMethod
	IDispatch createShortcut(String path);


    /**
     * WshShortcutインターフェイス.
     */
    public interface IWshShortcut {

        /**
         * ショートカットの実行可能ファイルへのパスを設定します。
         * @param path ショートカットの実行可能ファイルへのパス
         */
        @ComProperty
        void setTargetPath(String path);

        /**
         * ショートカット実行時の引数を設定します。
         * @param args ショートカット実行時の引数
         */
        @ComProperty
        void setArguments(String args);
        
        /**
         * ショートカットの作業ディレクトリを設定します。
         * @param dir 作業ディレクトリのパス
         */
        @ComProperty
        void setWorkingDirectory(String dir);
        
        /**
         * アイコンをショートカットに割り当てます。
         * 絶対パスと、アイコンに関連付けられているインデックスを含める必要があります。
         * @param path アイコンの格納場所を示すパス+インデックス
         */
        @ComProperty
        void setIconLocation(String path);
        
        /**
         * ショートカットに割り当てるキーの組み合わせを設定します。
         * @param keys キーの組み合わせを表す文字列
         */
        @ComProperty
        void setHotkey(String keys);
        
        /**
         * ウィンドウ スタイルをショートカットに割り当てます。
         * 1. ウィンドウをアクティブにして表示
         * 3. ウィンドウを最大化して表示
         * 7. ウィンドウを最小化して表示
         * @param style キーの組み合わせを表す文字列
         */
        @ComProperty
        void setWindowStyle(Integer style);
        
        /**
         * ショートカットの説明文（コメント）を設定します。
         * @param text 説明文（コメント）
         */
        @ComProperty
        void setDescription(String text);
        
        /**
         * ショートカットオブジェクトを保存します。
         */
        @ComMethod
        void save();
    }
}
