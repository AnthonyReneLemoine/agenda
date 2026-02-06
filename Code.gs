const SS_ID = "1kvCkyS-hkrP2c5JTx6X8QjV8J5G7n5IVgGYGLIbTt4o"; 
const SHEET_NAME = "Events";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle("Agenda")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Récupère tous les calendriers de l'utilisateur (hors masqués)
 */
function getUserCalendars() {
  const props = PropertiesService.getScriptProperties();
  const hiddenIds = JSON.parse(props.getProperty('HIDDEN_CAL_IDS') || '[]');
  
  return CalendarApp.getAllCalendars()
    .filter(cal => !hiddenIds.includes(cal.getId()))
    .map(cal => ({
      id: cal.getId(),
      name: cal.getName(),
      color: cal.getColor()
    }));
}

/**
 * Recherche un événement de manière robuste (gère les variations d'ID)
 * @param {Calendar} cal - Le calendrier
 * @param {string} eventId - L'ID de l'événement
 * @returns {CalendarEvent|null}
 */
function findEventById(cal, eventId) {
  if (!cal || !eventId) return null;
  
  // Nettoyer l'ID (parfois il y a des suffixes comme _R20250206)
  const cleanEventId = eventId.split('_')[0];
  
  // Essai 1: ID exact
  let event = cal.getEventById(eventId);
  if (event) return event;
  
  // Essai 2: ID nettoyé
  event = cal.getEventById(cleanEventId);
  if (event) return event;
  
  // Essai 3: Recherche dans les événements récents
  const now = new Date();
  const start = new Date(now);
  start.setDate(start.getDate() - 7);
  const end = new Date(now);
  end.setDate(end.getDate() + 60);
  
  const events = cal.getEvents(start, end);
  return events.find(e => {
    const eId = e.getId();
    return eId === eventId || eId === cleanEventId || eId.startsWith(cleanEventId);
  }) || null;
}

/**
 * Récupère les événements pour une liste de calendriers
 */
function getEventsForCalendars(calendarIds, viewStart, viewEnd) {
  if (!calendarIds || !calendarIds.length) return [];
  
  const start = new Date(viewStart);
  const end = new Date(viewEnd);

  const allEvents = [];

  for (const id of calendarIds) {
    try {
      const cal = CalendarApp.getCalendarById(id);
      if (!cal) continue;
      
      const calendarColor = cal.getColor();
      const events = cal.getEvents(start, end);
      
      for (const e of events) {
        const title = e.getTitle();
        
        allEvents.push({
          id: e.getId(),
          calendarId: id,
          title: title,
          start: e.getStartTime().toISOString(),
          end: e.getEndTime().toISOString(),
          allDay: e.isAllDayEvent(),
          backgroundColor: calendarColor,
          borderColor: calendarColor,
          textColor: '#ffffff',
          description: e.getDescription() || ''
        });
      }
    } catch (err) {
      console.error(`Erreur calendrier ${id}: ${err.message}`);
    }
  }
  
  return allEvents;
}

/**
 * Ajoute un ou plusieurs événements au calendrier (avec support multi-jours et récurrence)
 */
function addEventToCalendar(calendarId, eventData) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendrier introuvable');
  
  const options = { description: eventData.description || '' };
  let createdCount = 0;
  
  // Parser les dates
  const startDate = parseLocalDate(eventData.startDate);
  const endDate = parseLocalDate(eventData.endDate);
  
  // Déterminer si c'est un événement récurrent
  if (eventData.recurrence && eventData.recurrence !== 'none') {
    // Créer les événements récurrents
    const recurrenceEndDate = parseLocalDate(eventData.recurrenceEnd);
    let currentDate = new Date(startDate);
    
    while (currentDate <= recurrenceEndDate) {
      createSingleEvent(cal, eventData, currentDate, endDate, startDate, options);
      createdCount++;
      
      // Avancer selon le type de récurrence
      switch (eventData.recurrence) {
        case 'daily':
          currentDate.setDate(currentDate.getDate() + 1);
          break;
        case 'weekly':
          currentDate.setDate(currentDate.getDate() + 7);
          break;
        case 'biweekly':
          currentDate.setDate(currentDate.getDate() + 14);
          break;
        case 'monthly':
          currentDate.setMonth(currentDate.getMonth() + 1);
          break;
        default:
          currentDate = new Date(recurrenceEndDate.getTime() + 1); // Sortir de la boucle
      }
      
      // Sécurité : max 100 événements
      if (createdCount >= 100) break;
    }
  } else {
    // Événement unique
    createSingleEvent(cal, eventData, startDate, endDate, startDate, options);
    createdCount = 1;
  }
  
  logToSheet(eventData.title, eventData.startDate, eventData.endDate, eventData.description, cal.getName(), 'ADD');
  return { success: true, count: createdCount };
}

/**
 * Crée un seul événement
 */
function createSingleEvent(cal, eventData, currentStartDate, originalEndDate, originalStartDate, options) {
  if (eventData.allDay) {
    // Calculer le nombre de jours de l'événement original
    const daysDiff = Math.round((originalEndDate - originalStartDate) / (1000 * 60 * 60 * 24));
    
    if (daysDiff > 0) {
      // Événement multi-jours : endDate est exclusive dans createAllDayEvent
      const endDateExclusive = new Date(currentStartDate);
      endDateExclusive.setDate(endDateExclusive.getDate() + daysDiff + 1);
      cal.createAllDayEvent(eventData.title, currentStartDate, endDateExclusive, options);
    } else {
      // Événement d'une seule journée
      cal.createAllDayEvent(eventData.title, currentStartDate, options);
    }
  } else {
    // Événement avec heures
    const startDateTime = combineDateAndTime(currentStartDate, eventData.startTime);
    
    // Pour les événements multi-jours avec heures
    const daysDiff = Math.round((originalEndDate - originalStartDate) / (1000 * 60 * 60 * 24));
    const endDateForEvent = new Date(currentStartDate);
    endDateForEvent.setDate(endDateForEvent.getDate() + daysDiff);
    const endDateTime = combineDateAndTime(endDateForEvent, eventData.endTime);
    
    cal.createEvent(eventData.title, startDateTime, endDateTime, options);
  }
}

/**
 * Parse une date string en Date locale (sans décalage timezone)
 */
function parseLocalDate(dateString) {
  const parts = dateString.split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

/**
 * Combine une date et une heure en un seul objet Date
 */
function combineDateAndTime(date, timeString) {
  const [hours, minutes] = timeString.split(':').map(Number);
  const result = new Date(date);
  result.setHours(hours, minutes, 0, 0);
  return result;
}

/**
 * Met à jour un événement (avec possibilité de changement de calendrier)
 */
function updateEvent(oldCalId, newCalId, eventId, eventData) {
  const oldCal = CalendarApp.getCalendarById(oldCalId);
  if (!oldCal) throw new Error('Calendrier source introuvable');
  
  const isSameCalendar = (oldCalId === newCalId);
  const isAllDay = eventData.allDay === true;
  
  // Parser les dates
  const startDate = parseLocalDate(eventData.startDate);
  const endDate = parseLocalDate(eventData.endDate);
  
  if (isSameCalendar) {
    // Mise à jour simple sur le même calendrier
    const event = findEventById(oldCal, eventId);
    if (!event) throw new Error('Événement introuvable pour mise à jour (ID: ' + eventId + ')');
    
    event.setTitle(eventData.title);
    event.setDescription(eventData.description || '');
    
    if (isAllDay) {
      const daysDiff = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24));
      if (daysDiff > 0) {
        // Multi-jours
        const endDateExclusive = new Date(startDate);
        endDateExclusive.setDate(endDateExclusive.getDate() + daysDiff + 1);
        event.setAllDayDates(startDate, endDateExclusive);
      } else {
        event.setAllDayDate(startDate);
      }
    } else {
      const startDateTime = combineDateAndTime(startDate, eventData.startTime);
      const endDateTime = combineDateAndTime(endDate, eventData.endTime);
      event.setTime(startDateTime, endDateTime);
    }
    
    logToSheet(eventData.title, eventData.startDate, eventData.endDate, eventData.description, oldCal.getName(), 'UPDATE');
    return { success: true, eventId: event.getId(), calendarId: oldCalId };
    
  } else {
    // Déplacement vers un autre calendrier : CRÉER D'ABORD, SUPPRIMER ENSUITE
    const newCal = CalendarApp.getCalendarById(newCalId);
    if (!newCal) throw new Error('Calendrier destination introuvable');
    
    const options = { description: eventData.description || '' };
    let newEvent;
    
    if (isAllDay) {
      const daysDiff = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24));
      if (daysDiff > 0) {
        const endDateExclusive = new Date(startDate);
        endDateExclusive.setDate(endDateExclusive.getDate() + daysDiff + 1);
        newEvent = newCal.createAllDayEvent(eventData.title, startDate, endDateExclusive, options);
      } else {
        newEvent = newCal.createAllDayEvent(eventData.title, startDate, options);
      }
    } else {
      const startDateTime = combineDateAndTime(startDate, eventData.startTime);
      const endDateTime = combineDateAndTime(endDate, eventData.endTime);
      newEvent = newCal.createEvent(eventData.title, startDateTime, endDateTime, options);
    }
    
    if (!newEvent) {
      throw new Error('Échec de création du nouvel événement');
    }
    
    const newEventId = newEvent.getId();
    
    // SUPPRIMER l'ancien événement (seulement après création réussie)
    const oldEvent = findEventById(oldCal, eventId);
    if (oldEvent) {
      oldEvent.deleteEvent();
    }
    
    logToSheet(eventData.title, eventData.startDate, eventData.endDate, eventData.description, newCal.getName(), 'MOVE');
    return { success: true, eventId: newEventId, calendarId: newCalId };
  }
}

/**
 * Supprime un événement
 */
function deleteEvent(calendarId, eventId) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendrier introuvable');
  
  const event = findEventById(cal, eventId);
  if (!event) throw new Error('Événement introuvable (ID: ' + eventId + ')');
  
  const title = event.getTitle();
  event.deleteEvent();
  
  logToSheet(title, '', '', '', cal.getName(), 'DELETE');
  return true;
}

/**
 * Met à jour les horaires d'un événement (drag & drop)
 * Note: Ne pas utiliser pour les événements allDay
 */
function updateEventTimes(calendarId, eventId, newStart, newEnd) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) return false;
  
  const event = findEventById(cal, eventId);
  if (!event) return false;
  
  // Ne pas modifier les événements allDay via cette fonction
  if (event.isAllDayEvent()) {
    return false;
  }
  
  event.setTime(new Date(newStart), new Date(newEnd));
  return true;
}

/**
 * Réinitialise les calendriers masqués
 */
function resetHiddenCalendars() {
  PropertiesService.getScriptProperties().deleteProperty('HIDDEN_CAL_IDS');
  return true;
}

/**
 * Masque un calendrier de l'affichage
 */
function hideCalendar(calendarId) {
  const props = PropertiesService.getScriptProperties();
  const hiddenIds = JSON.parse(props.getProperty('HIDDEN_CAL_IDS') || '[]');
  
  if (!hiddenIds.includes(calendarId)) {
    hiddenIds.push(calendarId);
    props.setProperty('HIDDEN_CAL_IDS', JSON.stringify(hiddenIds));
  }
  return true;
}

/**
 * Journalise une action dans le spreadsheet
 */
function logToSheet(title, start, end, description, calendarName, action) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Titre', 'Début', 'Fin', 'Description', 'Calendrier', 'Action', 'Horodatage']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    
    sheet.appendRow([title, start, end, description || '', calendarName, action, new Date()]);
  } catch (err) {
    console.error(`Erreur log: ${err.message}`);
  }
}
