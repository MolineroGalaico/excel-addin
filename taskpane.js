Office.onReady(() => {
  document.getElementById('showDialog').addEventListener('click', () => {
    Office.context.ui.displayDialogAsync(
      'https://tuusuario.github.io/excel-addin/dialog.html', // Actualiza con tu URL real
      { height: 40, width: 30 },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
            console.log("Dato recibido:", args.message);
            dialog.close();
          });
        } else {
          console.error("Error al mostrar el diálogo:", asyncResult.error.message);
        }
      }
    );
  });
});
