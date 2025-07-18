function enviarCorreo() {
  const profesor = document.getElementById('profesor').value;
  const correoProfesor = document.getElementById('correoProfesor').value;
  const materia = document.getElementById('materia').value;
  const laboratorio = document.getElementById('laboratorio').value;
  const fecha = document.getElementById('fecha').value;
  const hora = document.getElementById('hora').value;
  const asunto = document.getElementById('asunto').value;
  const mensaje = document.getElementById('mensaje').value;

  if (!profesor || !correoProfesor || !materia || !laboratorio || !fecha || !hora || !asunto || !mensaje) {
    alert('Por favor completa todos los campos.');
    return;
  }

  const correos = [
    "ccuevas@uaslp.mx",
    "gerardo.ocana@uaslp.mx",
    "juan.ponce@uaslp.mx",
    "maricela.bravo@uaslp.mx",
    "brenda.campos@uaslp.mx",
    "secretaria.academica@fci.uaslp.mx"
  ];

  const cuerpoCorreo = `Profesor: ${profesor}
Correo del profesor: ${correoProfesor}
Materia: ${materia}
Laboratorio solicitado: ${laboratorio}
Fecha: ${fecha}
Hora: ${hora}

Mensaje:
${mensaje}`;

  const mailtoLink = `mailto:${correos.join(',')}` +
    `?cc=${correoProfesor}` +
    `&subject=${encodeURIComponent(asunto)}` +
    `&body=${encodeURIComponent(cuerpoCorreo)}`;

  window.location.href = mailtoLink;

  const wb = XLSX.utils.book_new();
  const ws_data = [
    ["Profesor", "Correo", "Materia", "Laboratorio", "Fecha", "Hora", "Asunto", "Mensaje"],
    [profesor, correoProfesor, materia, laboratorio, fecha, hora, asunto, mensaje]
  ];
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "Reservas");
  XLSX.writeFile(wb, `reserva_${profesor.replace(/\s+/g, '_')}_${fecha}.xlsx`);

  alert("¡Tu solicitud ha sido enviada con éxito! Se ha generado evidencia en Excel.");
}
