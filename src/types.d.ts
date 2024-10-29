// src/types.d.ts

export interface EventData {
  date: string;
  equipe: string;
  serviceType: string;
}

export type ServiceType = "12H Nuit" | "12H Jour" | "24H";

export interface ICalEvent {
  title: string;
  start: [number, number, number, number, number];
  end: [number, number, number, number, number];
  description: string;
}
