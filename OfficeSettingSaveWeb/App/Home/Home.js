/// <reference path="../App.js" />

(function () {
    "use strict";

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#button-document-save').click(documentSave);
            $('#button-document-read').click(documentRead);
            $('#button-document-reset').click(documentReset);
            $('#button-storage-save').click(storageSave);
            $('#button-storage-read').click(storageRead);
            $('#button-storage-reset').click(storageReset);
        });
    };

    // Officeドキュメントに設定値の操作 --------------------------------------------------
    // 保存
    function documentSave() {
        Office.context.document.settings.set('document-setting', $('#input-document').val());
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Officeドキュメントに設定値を保存することに失敗しました。');
            }
            else {
                write('Officeドキュメントに設定値を保存しました。');
            }
        });
    }

    // 読み込み
    function documentRead() {
        var val = Office.context.document.settings.get('document-setting');

        if (val == null || val == '') {
            write('Officeドキュメントに設定値は保存されていません。');
        }
        $('#input-document').val(val);
    }

    // リセット
    function documentReset() {
        Office.context.document.settings.remove('document-setting');
        $('#input-document').val('');
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Officeドキュメントに設定値をリセットすることに失敗しました。');
            }
            else {
                write('Officeドキュメントの設定値をリセットしました。');
            }
        });
    }


    // LoaclStorageに設定値の操作 --------------------------------------------------
    // 保存
    function storageSave() {
        window.localStorage.setItem('storage-setting', $('#input-storage').val());
        write('LocalStorageに設定値を保存しました。');
    }

    // 読み込み
    function storageRead() {
        var val = window.localStorage.getItem('storage-setting');

        if (val == null || val == '') {
            write('LocalStorageに設定値は保存されていません。');
        }
        $('#input-storage').val(val);
    }

    // リセット
    function storageReset() {
        window.localStorage.removeItem('storage-setting');
        $('#input-storage').val('');
        write('LocalStorageの設定値をリセットしました。');
    }


    function write(str) {
        app.showNotification(str);
    }
})();