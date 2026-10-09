export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { resource, event, data } = req.body;
  const WEBEX_TOKEN = process.env.WEBEX_BOT_TOKEN;
  const ESPACIO_SOPORTE_ID = process.env.WEBEX_SUPPORT_ROOM_ID;
  const GOOGLE_SHEET_URL = process.env.GOOGLE_SHEET_WEBHOOK_URL;

  try {
    // 1. EVENTO: El usuario le escribe al Bot (Crear Ticket)
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

      // Enviar la tarjeta al espacio de soporte técnico
      const ticketRes = await fetch('https://webexapis.com/v1/messages', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${WEBEX_TOKEN}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          roomId: ESPACIO_SOPORTE_ID,
          markdown: `Nueva incidencia de ${senderEmail}`,
          attachments: [
            {
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: {
                $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                type: 'AdaptiveCard',
                version: '1.2',
                body: [
                  {
                    type: 'TextBlock',
                    text: '🛠️ Nueva Incidencia Reportada',
                    weight: 'Bolder',
                    size: 'Medium',
                    color: 'Attention'
                  },
                  { type: 'TextBlock', text: `**FECHA DE REPORTE:** ${fechaActual}` },
                  { type: 'TextBlock', text: `**USUARIO:** ${senderEmail}` },
                  { type: 'TextBlock', text: `**FALLA REPORTADA:** ${textoIncidencia}`, wrap: true }
                ],
                actions: [
                  {
                    type: 'Action.Submit',
                    title: '🙋‍♂️ Tomar Ticket',
                    data: {
                      action: 'tomar_ticket',
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

      const ticketMessageData = await ticketRes.json();

      // Guardar ticket inicial en Google Sheets
      if (GOOGLE_SHEET_URL && ticketMessageData.id) {
        await fetch(GOOGLE_SHEET_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'crear_ticket',
            ticketId: ticketMessageData.id,
            fechaReporte: fechaActual,
            fallaReportada: textoIncidencia,
            usuario: senderEmail
          })
        }).catch(err => console.error('Error Google Sheets:', err));
      }

      return res.status(200).json({ status: 'ticket_created' });
    }

    // 2. EVENTO: Clics en botones (Tomar o Resolver Ticket)
    if (resource === 'attachmentActions' && event === 'created') {
      const actionRes = await fetch(`https://webexapis.com/v1/attachment/actions/${data.id}`, {
        headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
      });
      const actionData = await actionRes.json();

      const actionType = actionData.inputs?.action;
      const usuarioReporta = actionData.inputs?.usuarioReporta || 'Usuario';
      const fallaReportada = actionData.inputs?.fallaReportada || '';
      const fechaReporte = actionData.inputs?.fechaReporte || '';
      
      const tecnicoEmail = actionData.personId 
        ? await getEmailPersona(actionData.personId, WEBEX_TOKEN) 
        : 'Un técnico';

      const fechaAccion = new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' });

      // CASO 2A: EL TÉCNICO TOMA EL TICKET (DIAGNÓSTICO)
      if (actionType === 'tomar_ticket') {
        // Borrar tarjeta previa
        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        // Publicar nueva tarjeta con botón para "Resolver"
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${WEBEX_TOKEN}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `⏳ Incidencia de ${usuarioReporta} en proceso por ${tecnicoEmail}`,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: '⏳ Incidencia En Proceso', weight: 'Bolder', size: 'Medium', color: 'Warning' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**FALLA REPORTADA:** ${fallaReportada}`, wrap: true },
                    { type: 'TextBlock', text: `**DIAGNOSTICO:** ${fechaAccion} / ${tecnicoEmail}`, weight: 'Bolder' }
                  ],
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: '✅ Marcar como Resuelto',
                      data: {
                        action: 'resolver_ticket',
                        usuarioReporta: usuarioReporta,
                        fallaReportada: fallaReportada,
                        fechaReporte: fechaReporte,
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

        // Notificar al usuario por mensaje privado
        if (usuarioReporta) {
          await fetch('https://webexapis.com/v1/messages', {
            method: 'POST',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
              toPersonEmail: usuarioReporta,
              markdown: `🙋‍♂️ **Tu ticket ha sido asignado**\n\nEl técnico **${tecnicoEmail}** ha tomado tu caso en diagnóstico y se pondrá en contacto contigo.`
            })
          }).catch(() => {});
        }

        // Actualizar Google Sheets (Diagnóstico)
        if (GOOGLE_SHEET_URL && data.messageId) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'tomar_ticket',
              ticketId: data.messageId,
              diagnostico: `${fechaAccion} / ${tecnicoEmail}`,
              tecnico: tecnicoEmail
            })
          }).catch(() => {});
        }
      }

      // CASO 2B: EL TÉCNICO RESUELVE EL TICKET (CORRECCIÓN)
      if (actionType === 'resolver_ticket') {
        const tecnicoAtendio = actionData.inputs?.tecnicoAtendio || tecnicoEmail;
        const fechaDiagnostico = actionData.inputs?.fechaDiagnostico || fechaAccion;

        // Borrar tarjeta de "En Proceso"
        if (data.messageId) {
          await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
          }).catch(() => {});
        }

        // Publicar tarjeta final sin botones
        await fetch('https://webexapis.com/v1/messages', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${WEBEX_TOKEN}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            roomId: ESPACIO_SOPORTE_ID,
            markdown: `🎉 Incidencia de ${usuarioReporta} RESUELTA por ${tecnicoEmail}`,
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: {
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  type: 'AdaptiveCard',
                  version: '1.2',
                  body: [
                    { type: 'TextBlock', text: '🎉 Incidencia Resuelta', weight: 'Bolder', size: 'Medium', color: 'Good' },
                    { type: 'TextBlock', text: `**USUARIO:** ${usuarioReporta}` },
                    { type: 'TextBlock', text: `**DIAGNOSTICO:** ${fechaDiagnostico} / ${tecnicoAtendio}` },
                    { type: 'TextBlock', text: `**CORRECCIÓN:** ${fechaAccion} / ${tecnicoEmail}`, weight: 'Bolder' }
                  ]
                }
              }
            ]
          })
        });

        // Notificar al usuario que su caso fue resuelto
        if (usuarioReporta) {
          await fetch('https://webexapis.com/v1/messages', {
            method: 'POST',
            headers: { Authorization: `Bearer ${WEBEX_TOKEN}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
              toPersonEmail: usuarioReporta,
              markdown: `✅ **Tu ticket de soporte ha sido MARCADO COMO RESUELTO**\n\nEl técnico **${tecnicoEmail}** ha finalizado la atención de tu reporte. Si sigues experimentando problemas, por favor envía un nuevo mensaje.`
            })
          }).catch(() => {});
        }

        // Actualizar Google Sheets (Corrección y Cierre)
        if (GOOGLE_SHEET_URL && data.messageId) {
          await fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'resolver_ticket',
              ticketId: data.messageId,
              correccion: `${fechaAccion} / ${tecnicoEmail}`
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
