<script lang="ts">
  import ExcelJS from 'exceljs';
  import { createEvent, downloadICal } from './icalHelper';
  import type { EventData } from './types';

  let file: File | null = null;
  let sheets: string[] = [];
  let selectedSheetName: string | null = null;
  let names: string[] = [];
  let selectedName: string | null = null;
  let events: EventData[] = [];
  let allEvents: EventData[] = [];
  let errorMessage: string = '';

  const currentYear = new Date().getFullYear();

  /**
   * Gère la sélection d'un fichier Excel et charge les noms des feuilles valides.
   * @param event - L'événement de sélection de fichier
   */
  async function handleFile(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      file = input.files[0];
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);

      sheets = workbook.worksheets
        .map(sheet => sheet.name)
        .filter(sheetName => {
          const yearMatch = sheetName.match(/\b\d{4}\b/);
          return yearMatch && parseInt(yearMatch[0]) >= currentYear;
        });

      selectedSheetName = null;
      names = [];
      selectedName = null;
      errorMessage = '';
      console.log("Feuilles valides détectées :", sheets);
    }
  }

  /**
   * Gère la sélection de la feuille dans le fichier et extrait les événements et noms uniques.
   */
  async function handleSheetSelection() {
    if (!file || !selectedSheetName) {
      errorMessage = "Veuillez sélectionner un fichier et une feuille.";
      return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const worksheet = workbook.getWorksheet(selectedSheetName);

    if (worksheet) {
      allEvents = parseSheet(worksheet);
      extractUniqueNames(worksheet);
      errorMessage = '';
    } else {
      console.error("Erreur lors de la sélection de la feuille.");
    }
  }

  /**
   * Extrait les événements depuis la feuille Excel en capturant la couleur et le thème pour chaque cellule.
   * @param worksheet - La feuille Excel à analyser
   * @returns Liste des événements avec détails
   */
  function parseSheet(worksheet: ExcelJS.Worksheet): EventData[] {
    const extractedEvents: EventData[] = [];

    worksheet.eachRow((row, rowIndex) => {
      const rowValues = row.values.slice(1).map(value => value ? value.toString() : "");
      const date = rowValues[1];
      const equipe = rowValues[2];

      const extraData = rowValues.slice(3).map((value, index) => {
        const cell = row.getCell(index + 4);
        const color = cell.fill?.fgColor?.argb || null;
        const theme = cell.fill?.fgColor?.theme || null;

        return {
          value: value,
          color: color,
          color_theme: theme,
          serviceType: determineServiceType({ color, color_theme: theme })
        };
      });

      if (date && equipe) {
        extractedEvents.push({
          date,
          equipe,
          extraData
        });
      }
    });
    return extractedEvents;
  }

  /**
   * Détermine le type de garde en fonction de la couleur de la cellule.
   * @param cellData - Données de la cellule incluant la couleur et le thème
   * @returns Le type de garde : "12H Jour", "12H Nuit", ou "24H"
   */
  function determineServiceType(cellData: { color: string | null, color_theme: number | null }): string {
    if (cellData.color === 'FF00B0F0') {
      return '12H Jour';
    }
    if (cellData.color_theme === 9) {
      return '12H Nuit';
    }
    return '24H';
  }

/**
   * Extrait les noms uniques pour le menu déroulant en excluant les cellules non pertinentes.
   * @param worksheet - La feuille Excel à analyser pour les noms
   */
  function extractUniqueNames(worksheet: ExcelJS.Worksheet) {
      const uniqueNames = new Set<string>();
  
      let headerRow = worksheet.getRow(1);
      const isFirstRowEmpty = headerRow.values.every(cell => cell === null || cell === undefined);
  
      if (isFirstRowEmpty) {
          headerRow = worksheet.getRow(2);
      }
  
      worksheet.columns.slice(3).forEach((column, index) => {
          const header = headerRow.getCell(index + 4).value;
  
          if (header && header !== 'DESIDERATA') {
              column.eachCell({ includeEmpty: false }, (cell) => {
                  if (cell.value) { // Vérifie que cell.value est défini
                      const cellColor = cell.fill?.fgColor?.argb || null;
                      const cellTheme = cell.fill?.fgColor?.theme || null;
                      const cellValue = cell.value.toString();
                      const serviceType = determineServiceType({ color: cellColor, color_theme: cellTheme });
  
                      if (serviceType === '24H' && !/\d/.test(cellValue) && cellValue.trim() !== "") {
                          uniqueNames.add(cellValue);
                      }
                  }
              });
          }
      });
  
      names = Array.from(uniqueNames).sort((a, b) => a.localeCompare(b));
      selectedName = null;
  }

  /**
   * Filtre les événements en fonction du nom sélectionné.
   */
  function filterEvents() {
      if (selectedName) {
          events = allEvents.filter(event =>
              event.extraData.some(cellData => 
                  cellData.value && cellData.value.toUpperCase().trim() === selectedName.toUpperCase().trim()
              )
          );
      } else {
          events = [];
      }
      //console.log("Événements filtrés :", events);
  }

/**
   * Obtient le type de service pour un événement spécifique en fonction du nom sélectionné.
   * @param event - L'événement pour lequel obtenir le serviceType
   * @returns Le type de service correspondant à selectedName ou null si non trouvé
   */
  function getServiceTypeForSelectedName(event: EventData): string | null {
      const matchingCell = event.extraData.find(cell =>
          cell?.value?.toUpperCase().trim() === selectedName?.toUpperCase().trim()
      );
      return matchingCell ? matchingCell.serviceType : null;
  }

  /**
   * Formate la date pour un affichage en français avec le jour, mois et année.
   * @param dateString - La date en chaîne de caractères
   * @returns La date formatée
   */
  function formatDisplayDate(dateString: string): string {
    const date = new Date(dateString);
    return new Intl.DateTimeFormat('fr-FR', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(date);
  }

  /**
   * Génère un fichier iCal pour les événements filtrés.
   */
function generateICal() {
       const icalEvents = events.map(event => ({
           ...event,
           serviceType: getServiceTypeForSelectedName(event) || 'Inconnu'
       }));
       downloadICal(icalEvents); // Passe les événements enrichis à downloadICal
   }
</script>

<main>
  <h1>Gestion des Gardes</h1>

  <div class="form-container">
    <!-- Étape 1 : Sélection du fichier -->
    <label for="fileInput">Importer un fichier Excel :</label>
    <input id="fileInput" type="file" accept=".xlsx" on:change={handleFile} />

    <!-- Étape 2 : Sélection de la feuille si un fichier est sélectionné -->
    {#if sheets.length > 0}
      <label for="sheetSelect">Choisir une feuille :</label>
      <select id="sheetSelect" bind:value={selectedSheetName} on:change={handleSheetSelection}>
        <option value="" disabled selected={selectedSheetName === null}>Sélectionner une feuille</option>
        {#each sheets as sheet}
          <option value={sheet}>{sheet}</option>
        {/each}
      </select>
    {/if}

    <!-- Étape 3 : Sélection du nom unique si une feuille est sélectionnée -->
    {#if names.length > 0}
      <label for="nameSelect">Choisir un nom :</label>
      <select id="nameSelect" bind:value={selectedName} on:change={filterEvents}>
        <option value="" disabled selected={selectedName === null}>Sélectionner un nom</option>
        {#each names as name}
          <option value={name}>{name}</option>
        {/each}
      </select>
    {/if}

    <!-- Affichage du message d'erreur -->
    {#if errorMessage}
      <p class="error-message">{errorMessage}</p>
    {/if}
  </div>

  <!-- Affichage des événements filtrés -->
  {#if events.length > 0}
    <h2>Événements de Garde</h2>
    <ul class="event-list">
      {#each events as event}
        <li class="event-item">
          <span class="date">{formatDisplayDate(event.date)}</span>
          <span class="service-type">{getServiceTypeForSelectedName(event) || 'Inconnu'}</span>
          <span class="team">{event.equipe}</span>
        </li>
      {/each}
    </ul>
    <button class="button" on:click={generateICal}>Télécharger iCal</button>
  {:else if selectedName}
    <p>Aucun événement trouvé pour le nom sélectionné dans la feuille {selectedSheetName}</p>
  {/if}
</main>

<style>
  /* Ajoutez ici le style de votre choix pour structurer l'affichage */
</style>