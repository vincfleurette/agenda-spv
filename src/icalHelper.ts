// icalHelper.ts
import { createEvents } from "ics";
import type { EventData, ICalEvent } from "./types";

/**
 * Détermine les horaires de début et de fin en fonction du type de garde.
 * @param dateValue - La date de l'événement (string ou Date)
 * @param serviceType - Le type de garde : "12H Jour", "12H Nuit", ou "24H"
 * @returns Un objet contenant les dates de début et de fin de l'événement
 */
function getEventTimes(
  dateValue: string | Date,
  serviceType: string
): { start: Date; end: Date } {
  let date: Date;

  if (typeof dateValue === "string") {
    const parts = new Intl.DateTimeFormat("fr-FR", {
      day: "numeric",
      month: "long",
      year: "numeric",
    }).formatToParts(new Date(dateValue));

    const day = Number(parts.find((part) => part.type === "day")?.value);
    const month = new Date(
      Date.parse(parts.find((part) => part.type === "month")?.value + " 1")
    ).getMonth();
    const year = Number(parts.find((part) => part.type === "year")?.value);

    if (isNaN(day) || isNaN(month) || isNaN(year)) {
      throw new Error("Format de date invalide dans la chaîne : " + dateValue);
    }

    date = new Date(year, month, day);
  } else if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    throw new Error("Format de date non reconnu : " + dateValue);
  }

  const start = new Date(date);
  const end = new Date(start);

  switch (serviceType) {
    case "12H Jour":
      start.setHours(7, 30);
      end.setHours(19, 30);
      break;
    case "12H Nuit":
      start.setHours(19, 30);
      end.setDate(end.getDate() + 1);
      end.setHours(7, 30);
      break;
    case "24H":
      start.setHours(7, 30);
      end.setDate(end.getDate() + 1);
      end.setHours(7, 30);
      break;
    default:
      throw new Error("Type de service non reconnu : " + serviceType);
  }

  return { start, end };
}

/**
 * Crée un événement iCal à partir des données d'un événement.
 * @param eventData - Les données de l'événement contenant la date, l'équipe, et le type de service
 * @returns Un objet iCalEvent configuré pour la génération iCal
 */
export function createEvent({
  date,
  equipe,
  serviceType,
}: EventData): ICalEvent {
  const { start, end } = getEventTimes(date, serviceType);

  return {
    title: `Garde ${serviceType} - ${equipe}`,
    start: [
      start.getFullYear(),
      start.getMonth() + 1,
      start.getDate(),
      start.getHours(),
      start.getMinutes(),
    ],
    end: [
      end.getFullYear(),
      end.getMonth() + 1,
      end.getDate(),
      end.getHours(),
      end.getMinutes(),
    ],
    description: `Garde de ${serviceType} pour l'équipe ${equipe}`,
  };
}

/**
 * Génère et télécharge un fichier iCal (.ics) contenant tous les événements.
 * @param events - Liste des événements iCal enrichis avec `serviceType`
 */
export function downloadICal(events: EventData[]) {
  const icalEvents = events.map((event) => createEvent(event));
  createEvents(icalEvents, (error, value) => {
    if (error) {
      console.error("Erreur de création iCal :", error);
      return;
    }

    const blob = new Blob([value as BlobPart], { type: "text/calendar" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "garde_pompier.ics";
    a.click();
    URL.revokeObjectURL(url);
  });
}
