function selectedField(field) {
  var select = document.getElementById( 'fieldsSelector' );
  
  for ( var i = 0, l = select.options.length, o; i < l; i++ )
  {
    o = select.options[i];
    if ( field.id == o.value )
    {
      if (o.selected){
        o.selected = false;
        field.style.border = '1px solid transparent';
      }
      else{
        o.selected = true;
        field.style.border = '1px solid #022b4d';
      }
    }
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btn_load").onclick = load;
  }
});

// export async function load() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
