export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { resource, event, data } = req.body;
  const WEBEX_TOKEN = process.env.WEBEX_BOT_TOKEN;
  const ESPACIO_SOPORTE_ID = process.env.WEBEX_SUPPORT_ROOM_ID;

  try {
    // 1. EVENTO: El usuario le escribió al Bot en chat privado
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

      await fetch('https://webexapis.com/v1/messages', {
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
                  {
                    type: 'TextBlock',
                    text: `**Usuario:** ${senderEmail}`
                  },
                  {
                    type: 'TextBlock',
                    text: `**Detalle:** ${textoIncidencia}`,
                    wrap: true
                  }
                ],
                actions: [
                  {
                    type: 'Action.Submit',
                    title: '🙋‍♂️ Tomar Ticket',
                    data: {
                      action: 'tomar_ticket',
                      usuarioReporta: senderEmail
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

    // 2. EVENTO: Un técnico presionó "Tomar Ticket"
    if (resource === 'attachmentActions' && event === 'created') {
      // Obtener detalles de la acción
      const actionRes = await fetch(`https://webexapis.com/v1/attachment/actions/${data.id}`, {
        headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
      });
      const actionData = await actionRes.json();

      const tecnicoEmail = actionData.personId 
        ? await getEmailPersona(actionData.personId, WEBEX_TOKEN) 
        : 'Un técnico';
      const usuarioReporta = actionData.inputs?.usuarioReporta || 'Usuario';

      // 1. Borrar el mensaje original de la tarjeta con el botón
      if (data.messageId) {
        await fetch(`https://webexapis.com/v1/messages/${data.messageId}`, {
          method: 'DELETE',
          headers: { Authorization: `Bearer ${WEBEX_TOKEN}` }
        }).catch(() => {});
      }

      // 2. Publicar la tarjeta de confirmación de asignación
      await fetch('https://webexapis.com/v1/messages', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${WEBEX_TOKEN}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          roomId: ESPACIO_SOPORTE_ID,
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
                    text: '✅ Incidencia Asignada',
                    weight: 'Bolder',
                    size: 'Medium',
                    color: 'Good'
                  },
                  {
                    type: 'TextBlock',
                    text: `**Usuario:** ${usuarioReporta}`
                  },
                  {
                    type: 'TextBlock',
                    text: `**Atendido por:** ${tecnicoEmail}`,
                    weight: 'Bolder'
                  }
                ]
              }
            }
          ]
        })
      });

      return res.status(200).json({ status: 'ticket_assigned' });
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
    return data.emails?.[0] || 'Un técnico';
  } catch {
    return 'Un técnico';
  }
}
