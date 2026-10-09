export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { resource, event, data } = req.body;
  const WEBEX_TOKEN = process.env.WEBEX_BOT_TOKEN;
  const ESPACIO_SOPORTE_ID = process.env.WEBEX_SUPPORT_ROOM_ID;
  const GOOGLE_SHEET_URL = process.env.GOOGLE_SHEET_WEBHOOK_URL;

  try {
    // 1. CREAR TICKET (CUANDO EL USUARIO ESCRIBE AL BOT)
    if (resource === 'messages' && event === 'created') {
      const senderEmail = data.personEmail;

      if (senderEmail && senderEmail.endsWith('@webex.bot')) {
        return res.status(200).json({ status: 'ignored_bot' });
      }

      const msgRes = await fetch(`https://webexapis.com/v1/messages/${data.id}`, {
        headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
      });
      const msgData = await msgRes.json();
      const textoIncidencia = msgData.text || 'Sin detalle provisto';
      const fechaActual = new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' });

      let folioGenerado = 'PDS-TICKET';
      if (GOOGLE_SHEET_URL) {
        try {
          const sheetRes = await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'crear_ticket',
              fechaReporte: fechaActual,
              fallaReportada: textoIncidencia,
              usuario: senderEmail
            })
          });
          const sheetData = await sheetRes.json();
          if (sheetData.folio) folioGenerado = sheetData.folio;
        } catch (e) {
          console.error('Error enviando a Google Sheets:', e);
        }
      }

      await fetch('https://webexapis.com/v1/messages', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${WEBEX_TOKEN}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          roomId: ESPACIO_SOPORTE_ID,
          markdown: `Nueva incidencia [${folioGenerado}] de ${senderEmail}`,
          attachments: [
            {
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: {
                $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                type: 'AdaptiveCard',
                version: '1.2',
                body: [
                  { type: 'TextBlock', text: `🛠️ Ticket #${folioGenerado}`, weight: 'Bolder', size: 'Medium', color: 'Attention' },
                  { type: 'TextBlock', text: `**FECHA DE REPORTE:** ${fechaActual}` },
                  { type: 'TextBlock', text: `**USUARIO:** ${senderEmail}` },
                  { type: 'TextBlock', text: `**FALLA REPORTADA:** ${textoIncidencia}`, wrap: true },
                  { type: 'Input.Text', id: 'inputDiagnostico', placeholder: 'Escribe el diagnóstico inicial...', isMultiline: true }
                ],
                actions: [
                  {
                    type: 'Action.Submit',
                    title: '🙋‍♂️ Tomar Ticket',
                    data: {
                      action: 'tomar_ticket',
                      folio: folioGenerado,
                      usuarioReporta: senderEmail,
                      fallaReportada: textoIncidencia,
                      fechaReporte: fechaActual
                    }
                  }
                ]
              }
            }
          ]
        })
      });

      return res.status(200).json({ status: 'ticket_created' });
    }

    // 2. ACCIONES DE LAS TARJETAS (SUBMIT)
    if (resource === 'attachmentActions' && event === 'created') {
      const actionRes = await fetch(`https://webexapis.com/v1/attachment/actions/${data.id}`, {
        headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
      });
      const actionData = await actionRes.json();

      const inputs = actionData.inputs || {};
      const actionType = inputs.action;
      const folio = inputs.folio || 'FOLIO';
      const usuarioReporta = inputs.usuarioReporta || 'Usuario';
      const fallaReportada = inputs.fallaReportada || '';
      const fechaReporte = inputs.fechaReporte || '';
      
      const tecnicoEmail = actionData.personId 
        ? await getEmailPersona(actionData.personId, WEBEX_TOKEN) 
        : 'Un técnico';

      const fechaAccion = new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' });

      // CASO A: EL TÉCNICO TOMA EL TICKET (SIN ENVIAR DIAGNÓSTICO AL USUARIO)
      if (actionType === 'tomar_ticket') {
        const textoDiagnostico = inputs.inputDiagnostico || 'En revisión por técnico';

        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${WEBEX_TOKEN}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `⏳ Ticket #${folio} en proceso por ${tecnicoEmail}`,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: `⏳ Ticket #${folio} En Proceso`, weight: 'Bolder', size: 'Medium', color: 'Warning' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**FALLA REPORTADA:** ${fallaReportada}`, wrap: true },
                    { type: 'TextBlock', text: `**DIAGNOSTICO:** ${textoDiagnostico}`, wrap: true },
                    { type: 'TextBlock', text: `**FECHA / RESPONSABLE:** ${fechaAccion} / ${tecnicoEmail}`, weight: 'Bolder' },
                    { type: 'Input.Text', id: 'inputCorreccion', placeholder: 'Escribe la solución realizada...', isMultiline: true }
                  ],
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: '✅ Marcar como Resuelto',
                      data: {
                        action: 'resolver_ticket',
                        folio: folio,
                        usuarioReporta: usuarioReporta,
                        fallaReportada: fallaReportada,
                        fechaReporte: fechaReporte,
                        diagnosticoPrevio: textoDiagnostico,
                        tecnicoAtendio: tecnicoEmail,
                        fechaDiagnostico: fechaAccion
                      }
                    }
                  ]
                }
              }
            ]
          })
        });

        // Notificar al usuario (solo que fue asignado)
        if (usuarioReporta) {
          await fetch('https://webexapis.com/v1/messages', {
            method: 'POST',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
              toPersonEmail: usuarioReporta,
              markdown: `🙋‍♂️ **Ticket #${folio} Asignado**\n\nEl técnico **${tecnicoEmail}** ha tomado tu caso y se pondrá en contacto contigo.`
            })
          }).catch(() => {});
        }

        if (GOOGLE_SHEET_URL) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'tomar_ticket',
              folio: folio,
              diagnostico: textoDiagnostico,
              fecha: fechaAccion,
              tecnico: tecnicoEmail
            })
          }).catch(() => {});
        }
      }

      // CASO B: EL TÉCNICO RESUELVE EL TICKET -> ENVÍA TARJETA INTERACTIVA DE CONFIRMACIÓN AL USUARIO
      if (actionType === 'resolver_ticket') {
        const textoCorreccion = inputs.inputCorreccion || 'Problema corregido';
        const diagnosticoPrevio = inputs.diagnosticoPrevio || '';

        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${WEBEX_TOKEN}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `⏳ Ticket #${folio} en espera de confirmación del usuario`, wrap: true,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: `⏳ Ticket #${folio} Atendido (Esperando Confirmación)`, weight: 'Bolder', size: 'Medium', color: 'Warning' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**DIAGNOSTICO:** ${diagnosticoPrevio}` },
                    { type: 'TextBlock', text: `**CORRECCIÓN:** ${textoCorreccion}`, wrap: true },
                    { type: 'TextBlock', text: `**FECHA / RESPONSABLE:** ${fechaAccion} / ${tecnicoEmail}`, weight: 'Bolder' }
                  ]
                }
              }
            ]
          })
        });

        // Enviar tarjeta interactiva al usuario con los 2 botones de confirmación
        if (usuarioReporta) {
          await fetch('https://webexapis.com/v1/messages', {
            method: 'POST',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
              toPersonEmail: usuarioReporta,
              markdown: `¿Se solucionó el problema del Ticket #${folio}?`,
              attachments: [
                {
                  contentType: 'application/vnd.microsoft.card.adaptive',
                  content: {
                    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                    type: 'AdaptiveCard',
                    version: '1.2',
                    body: [
                      { type: 'TextBlock', text: `❓ Confirmación de Ticket #${folio}`, weight: 'Bolder', size: 'Medium' },
                      { 
                        type: 'TextBlock', 
                        text: `El técnico **${tecnicoEmail}** indicó la siguiente solución:\n\n*${textoCorreccion}*`,
                        wrap: true 
                      },
                      { type: 'TextBlock', text: '¿Confirmas que el problema quedó resuelto?', wrap: true }
                    ],
                    actions: [
                      {
                        type: 'Action.Submit',
                        title: '👍 Sí, quedó resuelto',
                        data: {
                          action: 'confirmar_usuario',
                          folio: folio,
                          usuarioReporta: usuarioReporta,
                          tecnicoEmail: tecnicoEmail,
                          textoCorreccion: textoCorreccion
                        }
                      },
                      {
                        type: 'Action.Submit',
                        title: '👎 No, sigo con la falla',
                        data: {
                          action: 'rechazar_usuario',
                          folio: folio,
                          usuarioReporta: usuarioReporta,
                          tecnicoEmail: tecnicoEmail,
                          fallaReportada: fallaReportada
                        }
                      }
                    ]
                  }
                }
              ]
            })
          }).catch(() => {});
        }

        if (GOOGLE_SHEET_URL) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'resolver_ticket',
              folio: folio,
              correccion: textoCorreccion,
              fecha: fechaAccion,
              tecnico: tecnicoEmail
            })
          }).catch(() => {});
        }
      }

      // CASO C: EL USUARIO CONFIRMA ("SÍ, QUEDÓ RESUELTO")
      if (actionType === 'confirmar_usuario') {
        // Elimina la tarjeta de confirmación enviada al usuario
        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        // Mensaje privado final al usuario
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            toPersonEmail: usuarioReporta,
            markdown: `🎉 **Ticket #${folio} Resuelto**\n\nGracias por confirmar. El ticket ha sido cerrado con éxito.`
          })
        }).catch(() => {});

        // Actualizar tarjeta en el Espacio de Soporte a VERDE (Resuelto y Confirmado)
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `🎉 Ticket #${folio} RESUELTO Y CONFIRMADO POR USUARIO`,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: `🎉 Ticket #${folio} Resuelto y Confirmado`, weight: 'Bolder', size: 'Medium', color: 'Good' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**ESTADO:** Confirmado por el usuario (${fechaAccion})` }
                  ]
                }
              }
            ]
          })
        });

        if (GOOGLE_SHEET_URL) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'confirmar_usuario',
              folio: folio,
              fecha: fechaAccion
            })
          }).catch(() => {});
        }
      }

      // CASO D: EL USUARIO RECHAZA ("NO, SIGO CON LA FALLA") -> VUELVE A PONERLO EN PENDIENTE EN LA SALA
      if (actionType === 'rechazar_usuario') {
        // Elimina la tarjeta de confirmación del usuario
        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        // Mensaje privado al usuario
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            toPersonEmail: usuarioReporta,
            markdown: `⚠️ **Ticket #${folio} Reabierto**\n\nHemos notificado al área de soporte que la falla persiste.`
          })
        }).catch(() => {});

        // Reabrir la tarjeta en el Espacio de Soporte con botón para volver a tomarlo
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `⚠️ Ticket #${folio} REABIERTO por el usuario`,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: `⚠️ Ticket #${folio} Reabierto`, weight: 'Bolder', size: 'Medium', color: 'Attention' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**FALLA REPORTADA:** ${fallaReportada}` },
                    { type: 'TextBlock', text: `*El usuario indicó que NO se resolvió el problema.*` },
                    { type: 'Input.Text', id: 'inputDiagnostico', placeholder: 'Escribe el nuevo diagnóstico...', isMultiline: true }
                  ],
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: '🙋‍♂️ Retomar Ticket',
                      data: {
                        action: 'tomar_ticket',
                        folio: folio,
                        usuarioReporta: usuarioReporta,
                        fallaReportada: fallaReportada
                      }
                    }
                  ]
                }
              }
            ]
          })
        });

        if (GOOGLE_SHEET_URL) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'rechazar_usuario',
              folio: folio,
              fecha: fechaAccion
            })
          }).catch(() => {});
        }
      }

      return res.status(200).json({ status: 'action_completed' });
    }

    return res.status(200).json({ status: 'ok' });
  } catch (error) {
    console.error('Error procesando webhook:', error);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
}

async function getEmailPersona(personId, token) {
  try {
    const res = await fetch(`https://webexapis.com/v1/people/${personId}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const data = await res.json();
    return data.emails?.[0] || data.displayName || 'Un técnico';
  } catch {
    return 'Un técnico';
  }
}
