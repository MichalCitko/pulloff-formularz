import { useEffect, useState } from 'react'
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import './App.css'

type BreakType = 'Klej' | 'Farba' | 'Tynk' | 'Zaprawa' | 'Inne'

type Measurement = {
  id: number
  point: string
  diameter: string
  location: string
  photo: string
  photoName: string
  resultPhoto: string
  resultPhotoName: string
  additionalPhoto: string
  additionalPhotoName: string
  breakType: BreakType
  value: string
  notes: string
}

type FormState = {
  objectName: string
  surveyAuthor: string
  participants: string
  structureType: string
  structureCondition: string
  floorsCount: string
  protectionDescription: string
}

const initialForm: FormState = {
  objectName: '',
  surveyAuthor: '',
  participants: '',
  structureType: '',
  structureCondition: '',
  floorsCount: '',
  protectionDescription: '',
}

const initialMeasurements: Measurement[] = [
  {
    id: 1,
    point: '1',
    diameter: '50',
    location: 'Ściana elewacyjna, strefa południowa, 1. kondygnacja',
    photo: '',
    photoName: '',
    resultPhoto: '',
    resultPhotoName: '',
    additionalPhoto: '',
    additionalPhotoName: '',
    breakType: 'Klej',
    value: '',
    notes: '',
  },
]

const diameterOptions = ['20', '50']
const breakOptions: BreakType[] = ['Klej', 'Farba', 'Tynk', 'Zaprawa', 'Inne']
const constructionOptions = [
  'Konstrukcja stalowa',
  'Konstrukcja betonowa',
  'Konstrukcja stalowa i betonowa',
]
const steelConditionOptions = [
  'Powłoka antykorozyjna w stanie dobrym bez widocznej korozji',
  'Powłoka antykorozyjna w stanie dobrym z widoczną korozją',
  'Powłoka antykorozyjna w stanie złym',
]
const concreteConditionOptions = [
  'Podłoże bez warstwy tynku',
  'Podłoże z warstwą tynku',
  'Podłoże pokryte farbą',
]

const storageKey = 'pulloff-report-v1'
const maxMeasurementCount = 15
const maxPhotoDimension = 800

const toBase64 = (dataUrl: string) => dataUrl.split(',')[1] ?? ''

const compressPhoto = async (file: File) => {
  const bitmap = await createImageBitmap(file)
  const scale = Math.min(1, maxPhotoDimension / Math.max(bitmap.width, bitmap.height))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(bitmap.width * scale))
  canvas.height = Math.max(1, Math.round(bitmap.height * scale))
  const context = canvas.getContext('2d')

  if (!context) {
    bitmap.close()
    throw new Error('Nie można przetworzyć zdjęcia.')
  }

  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
  bitmap.close()

  const blob = await new Promise<Blob | null>((resolve) => {
    canvas.toBlob(resolve, 'image/jpeg', 0.55)
  })

  if (!blob) throw new Error('Nie można skompresować zdjęcia.')

  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result ?? ''))
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })
}

const createCell = (text: string, isHeader = false) =>
  new TableCell({
    shading: isHeader
      ? {
          fill: 'D9EAF7',
          color: '000000',
        }
      : undefined,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: 'C0C0C0' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'C0C0C0' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'C0C0C0' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'C0C0C0' },
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text, bold: isHeader, size: isHeader ? 20 : 18 })],
      }),
    ],
  })

function App() {
  const [form, setForm] = useState<FormState>(() => {
    const stored = localStorage.getItem(storageKey)
    if (!stored) return initialForm

    try {
      const parsed = JSON.parse(stored) as { form?: FormState; measurements?: Measurement[] }
      return parsed.form ?? initialForm
    } catch {
      return initialForm
    }
  })

  const [measurements, setMeasurements] = useState<Measurement[]>(() => {
    const stored = localStorage.getItem(storageKey)
    if (!stored) return initialMeasurements

    try {
      const parsed = JSON.parse(stored) as { form?: FormState; measurements?: Measurement[] }
      return parsed.measurements?.length ? parsed.measurements : initialMeasurements
    } catch {
      return initialMeasurements
    }
  })

  const [status, setStatus] = useState('Dane zapisują się lokalnie na urządzeniu.')

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify({ form, measurements }))
  }, [form, measurements])

  const updateField = (field: keyof FormState, value: string) => {
    setForm((current) => ({ ...current, [field]: value }))
  }

  const addMeasurement = () => {
    setMeasurements((current) => {
      if (current.length >= maxMeasurementCount) return current

      return [
        ...current,
        {
        id: Date.now(),
        point: String(current.length + 1),
        diameter: '50',
        location: '',
        photo: '',
        photoName: '',
        resultPhoto: '',
        resultPhotoName: '',
        additionalPhoto: '',
        additionalPhotoName: '',
        breakType: 'Klej',
        value: '',
        notes: '',
        },
      ]
    })
  }

  const updateMeasurement = <K extends keyof Measurement>(
    id: number,
    field: K,
    value: Measurement[K],
  ) => {
    setMeasurements((current) =>
      current.map((item) => (item.id === id ? { ...item, [field]: value } : item)),
    )
  }

  const removeMeasurement = (id: number) => {
    setMeasurements((current) => {
      if (current.length === 1) {
        return current
      }
      return current.filter((item) => item.id !== id)
    })
  }

  const handlePhotoUpload = async (
    id: number,
    event: React.ChangeEvent<HTMLInputElement>,
    photoField: 'photo' | 'resultPhoto' | 'additionalPhoto',
    photoNameField: 'photoName' | 'resultPhotoName' | 'additionalPhotoName',
  ) => {
    const file = event.target.files?.[0]
    if (!file) return

    setStatus('Zmniejszanie zdjęcia...')

    try {
      const result = await compressPhoto(file)
      updateMeasurement(id, photoField, result)
      updateMeasurement(id, photoNameField, file.name)
      setStatus('Zdjęcie zostało zmniejszone i zapisane.')
    } catch {
      setStatus('Nie udało się przetworzyć zdjęcia. Spróbuj wybrać inny plik.')
    }
  }

  const generateWordReport = async () => {
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({
              text: 'PROTOKÓŁ BADANIA PULL-OFF',
              heading: HeadingLevel.TITLE,
              spacing: { after: 240 },
            }),
            new Paragraph({
              text: 'Aplikacja PULLOFF',
              heading: HeadingLevel.HEADING_1,
              spacing: { after: 180 },
            }),
            new Paragraph({ text: `Nazwa obiektu: ${form.objectName || '-'}` }),
            new Paragraph({ text: `Badanie prowadzi: ${form.surveyAuthor || '-'}` }),
            new Paragraph({ text: `Uczestnicy: ${form.participants || '-'}` }),
            new Paragraph({ text: `Konstrukcja: ${form.structureType || '-'}` }),
            new Paragraph({ text: `Stan konstrukcji: ${form.structureCondition || '-'}` }),
            new Paragraph({ text: `Liczba kondygnacji: ${form.floorsCount || '-'}` }),
            new Paragraph({ text: `Opis zabezpieczenia: ${form.protectionDescription || '-'}` }),
            new Paragraph({
              text: 'Punkty pomiarowe',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 240, after: 120 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    createCell('Lp.', true),
                    createCell('Średnica', true),
                    createCell('Miejsce', true),
                    createCell('Miejsce zerwania', true),
                    createCell('Wartość', true),
                    createCell('Uwagi', true),
                  ],
                }),
                ...measurements.map((item, index) =>
                  new TableRow({
                    children: [
                      createCell(String(index + 1)),
                      createCell(`${item.diameter} mm`),
                      createCell(item.location || '-'),
                      createCell(item.breakType),
                      createCell(item.value ? `${item.value} MPa` : '-'),
                      createCell(item.notes || '-'),
                    ],
                  }),
                ),
              ],
            }),
            new Paragraph({
              text: 'Zdjęcia punktów pomiarowych',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 240, after: 120 },
            }),
            ...measurements.flatMap((item, index) => {
              const imageParagraph: Paragraph[] = [
                new Paragraph({
                  spacing: { after: 120 },
                  children: [
                    new TextRun({
                      text: `Punkt ${index + 1} — ${item.location || 'brak opisu'} `,
                      bold: true,
                    }),
                  ],
                }),
              ]

              const photos = [
                {
                  label: 'Zdjęcie - punkt pomiarowy',
                  data: item.photo,
                  fileName: item.photoName,
                },
                {
                  label: 'Zdjęcie wyniku pomiaru',
                  data: item.resultPhoto,
                  fileName: item.resultPhotoName,
                },
                {
                  label: 'Zdjęcie dodatkowe',
                  data: item.additionalPhoto,
                  fileName: item.additionalPhotoName,
                },
              ]

              photos.forEach(({ label, data, fileName }) => {
                imageParagraph.push(
                  new Paragraph({
                    spacing: { before: 120 },
                    children: [new TextRun({ text: label, bold: true })],
                  }),
                )

                if (data) {
                  const imageData = Uint8Array.from(atob(toBase64(data)), (char) =>
                    char.charCodeAt(0),
                  )

                  imageParagraph.push(
                    new Paragraph({
                      children: [
                        new ImageRun({
                          data: imageData,
                          transformation: { width: 280, height: 180 },
                          type: 'jpg',
                        }),
                      ],
                    }),
                  )
                }

                imageParagraph.push(
                  new Paragraph({
                    spacing: { after: 120 },
                    children: [
                      new TextRun({ text: `Nazwa pliku: ${fileName || 'brak zdjęcia'}` }),
                    ],
                  }),
                )
              })

              return imageParagraph
            }),
            new Paragraph({
              text: 'Wnioski:',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 240, after: 120 },
            }),
            new Paragraph({
              text: 'Minimalna wymagana przyczepność natrysku ogniochronnego mcr TECWOOL F wynosi:',
              spacing: { after: 80 },
            }),
            new Paragraph({
              text: '• dla podłoża betonowego: 0,30 MPa (dopuszczalne minimum 80% wartości wymaganej: 0,24 MPa),',
              spacing: { after: 60 },
            }),
            new Paragraph({
              text: '• dla podłoża stalowego ze znaną powłoką antykorozyjną: 2,50 MPa (dopuszczalne minimum 80% wartości wymaganej: 2,00 MPa),',
              spacing: { after: 60 },
            }),
            new Paragraph({
              text: '• dla podłoża stalowego z nieznaną powłoką antykorozyjną: 5,00 MPa.',
              spacing: { after: 180 },
            }),
          ],
        },
      ],
    })

    const blob = await Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `${(form.objectName || 'protokol-pulloff').replace(/\s+/g, '-').toLowerCase()}.docx`
    link.click()
    URL.revokeObjectURL(url)
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Aplikacja PULLOFF</p>
          <h1>Protokół badania pull-off</h1>
        </div>
      </header>

      <p className="status-line">{status}</p>

      <section className="card">
        <h2>Dane obiektu i badania</h2>
        <div className="form-grid">
          <label>
            <span>Nazwa obiektu</span>
            <input
              value={form.objectName}
              onChange={(event) => updateField('objectName', event.target.value)}
              placeholder="np. Budynek przy ul. ..."
            />
          </label>

          <label>
            <span>Badanie prowadzi</span>
            <input
              value={form.surveyAuthor}
              onChange={(event) => updateField('surveyAuthor', event.target.value)}
              placeholder="Imię i nazwisko / firma"
            />
          </label>

          <label className="full-width">
            <span>Osoba / osoby przy udziale</span>
            <input
              value={form.participants}
              onChange={(event) => updateField('participants', event.target.value)}
              placeholder="np. kierownik budowy, inspektor, inwestor"
            />
          </label>

          <label>
            <span>Konstrukcja</span>
            <select
              value={form.structureType}
              onChange={(event) => {
                const structureType = event.target.value
                const conditionOptions =
                  structureType === 'Konstrukcja stalowa'
                    ? steelConditionOptions
                    : structureType === 'Konstrukcja betonowa'
                      ? concreteConditionOptions
                      : []
                setForm((current) => ({
                  ...current,
                  structureType,
                  structureCondition:
                    conditionOptions.length > 0 &&
                    !conditionOptions.includes(current.structureCondition)
                      ? ''
                      : current.structureCondition,
                }))
              }}
            >
              <option value="">Wybierz konstrukcję</option>
              {constructionOptions.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </label>

          <label>
            <span>Stan konstrukcji</span>
            {form.structureType === 'Konstrukcja stalowa' ||
            form.structureType === 'Konstrukcja betonowa' ? (
              <select
                value={form.structureCondition}
                onChange={(event) => updateField('structureCondition', event.target.value)}
              >
                <option value="">Wybierz stan konstrukcji</option>
                {(form.structureType === 'Konstrukcja stalowa'
                  ? steelConditionOptions
                  : concreteConditionOptions
                ).map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            ) : (
              <input
                value={form.structureCondition}
                onChange={(event) => updateField('structureCondition', event.target.value)}
                placeholder="np. dobry, wymagający naprawy, wilgotny"
              />
            )}
          </label>

          <label>
            <span>Liczba kondygnacji</span>
            <input
              value={form.floorsCount}
              onChange={(event) => updateField('floorsCount', event.target.value)}
              placeholder="np. 5"
            />
          </label>

          <label className="full-width">
            <span>Opis zabezpieczenia / warstwa</span>
            <textarea
              value={form.protectionDescription}
              onChange={(event) => updateField('protectionDescription', event.target.value)}
              placeholder="np. płyta EPS, tynk mineralny, farba akrylowa"
            />
          </label>

        </div>
        <div className="form-actions">
          <button type="button" className="primary-button" onClick={generateWordReport}>
            Generuj Word
          </button>
        </div>
      </section>

      <section className="card">
        <div className="section-header">
          <h2>Punkty pomiarowe</h2>
          <button
            type="button"
            className="secondary-button"
            onClick={addMeasurement}
            disabled={measurements.length >= maxMeasurementCount}
          >
            + Dodaj punkt ({measurements.length}/{maxMeasurementCount})
          </button>
        </div>

        <div className="measurement-list">
          {measurements.map((measurement, index) => (
            <article key={measurement.id} className="measurement-card">
              <div className="measurement-topline">
                <h3>Punkt pomiarowy {index + 1}</h3>
                {measurements.length > 1 && (
                  <button
                    type="button"
                    className="delete-button"
                    onClick={() => removeMeasurement(measurement.id)}
                  >
                    Usuń
                  </button>
                )}
              </div>

              <div className="measurement-grid">
                <label>
                  <span>Nr punktu</span>
                  <input
                    value={measurement.point}
                    onChange={(event) => updateMeasurement(measurement.id, 'point', event.target.value)}
                  />
                </label>

                <label>
                  <span>Średnica grzybka</span>
                  <select
                    value={measurement.diameter}
                    onChange={(event) =>
                      updateMeasurement(measurement.id, 'diameter', event.target.value)
                    }
                  >
                    {diameterOptions.map((option) => (
                      <option key={option} value={option}>
                        {option} mm
                      </option>
                    ))}
                  </select>
                </label>

                <label className="full-width">
                  <span>Miejsce pomiaru</span>
                  <input
                    value={measurement.location}
                    onChange={(event) => updateMeasurement(measurement.id, 'location', event.target.value)}
                    placeholder="np. Ściana elewacyjna, lewa strona, 2. kondygnacja"
                  />
                </label>

                <label>
                  <span>Wartość pomiaru</span>
                  <input
                    value={measurement.value}
                    onChange={(event) => updateMeasurement(measurement.id, 'value', event.target.value)}
                    placeholder="np. 1,25"
                  />
                </label>

                <label>
                  <span>Miejsce zerwania</span>
                  <select
                    value={measurement.breakType}
                    onChange={(event) =>
                      updateMeasurement(measurement.id, 'breakType', event.target.value as BreakType)
                    }
                  >
                    {breakOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="full-width">
                  <span>Uwagi</span>
                  <textarea
                    value={measurement.notes}
                    onChange={(event) => updateMeasurement(measurement.id, 'notes', event.target.value)}
                    placeholder="Dodatkowe informacje o przyczepności lub uszkodzeniu"
                  />
                </label>

                <div className="photo-box full-width">
                  {[
                    {
                      label: 'Zdjęcie - punkt pomiarowy',
                      field: 'photo',
                      nameField: 'photoName',
                      data: measurement.photo,
                      fileName: measurement.photoName,
                      alt: `Zdjęcie punktu pomiarowego ${index + 1}`,
                    },
                    {
                      label: 'Zdjęcie wyniku pomiaru',
                      field: 'resultPhoto',
                      nameField: 'resultPhotoName',
                      data: measurement.resultPhoto,
                      fileName: measurement.resultPhotoName,
                      alt: `Zdjęcie wyniku pomiaru ${index + 1}`,
                    },
                    {
                      label: 'Zdjęcie dodatkowe',
                      field: 'additionalPhoto',
                      nameField: 'additionalPhotoName',
                      data: measurement.additionalPhoto,
                      fileName: measurement.additionalPhotoName,
                      alt: `Zdjęcie dodatkowe ${index + 1}`,
                    },
                  ].map((photo) => (
                    <div className="photo-group" key={photo.field}>
                      <label className="photo-upload">
                        <span>{photo.label}</span>
                        <input
                          type="file"
                          accept="image/*"
                          capture="environment"
                          onChange={(event) =>
                            handlePhotoUpload(
                              measurement.id,
                              event,
                              photo.field as 'photo' | 'resultPhoto' | 'additionalPhoto',
                              photo.nameField as 'photoName' | 'resultPhotoName' | 'additionalPhotoName',
                            )
                          }
                        />
                      </label>

                      {photo.data && (
                        <div className="photo-preview">
                          <img src={photo.data} alt={photo.alt} />
                          <small>{photo.fileName || 'Zdjęcie dodane'}</small>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            </article>
          ))}
        </div>
      </section>
    </div>
  )
}

export default App
