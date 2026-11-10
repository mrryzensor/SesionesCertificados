'use client';

import React, { useState, useEffect } from 'react';
import { 
  Sparkles, BookOpen, Download, FileText, Trash2, Edit3, Save, 
  HelpCircle, CheckCircle, Award, ListChecks, ChevronRight, Layers, Eye
} from 'lucide-react';
import { SessionPlan } from '@/lib/types';
import { useToast } from '@/components/Toast';
import { Modal } from '@/components/Modal';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, ShadingType } from 'docx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

export function SessionGenerator() {
  const { success, error, info } = useToast();

  // State
  const [theme, setTheme] = useState('');
  const [sourceContent, setSourceContent] = useState('');
  const [area, setArea] = useState('Desarrollo Personal, Ciudadanía y Cívica');
  const [grade, setGrade] = useState('3ro de Secundaria');
  const [section, setSection] = useState('A');
  const [teacherName, setTeacherName] = useState('Prof. David Mendoza');
  const [duration, setDuration] = useState('90 minutos');

  const [generating, setGenerating] = useState(false);
  const [sessions, setSessions] = useState<SessionPlan[]>([]);
  const [activeSession, setActiveSession] = useState<SessionPlan | null>(null);

  // Student test view modal
  const [isTestModalOpen, setIsTestModalOpen] = useState(false);
  const [studentAnswers, setStudentAnswers] = useState<Record<string, string>>({});
  const [testScore, setTestScore] = useState<{ total: number; max: number } | null>(null);

  useEffect(() => {
    loadSessions();
  }, []);

  const loadSessions = async () => {
    try {
      const res = await fetch('/api/sessions');
      const data = await res.json();
      if (data.sessions) {
        setSessions(data.sessions);
        if (data.sessions.length > 0 && !activeSession) {
          setActiveSession(data.sessions[0]);
        }
      }
    } catch (e) {
      console.error(e);
    }
  };

  const handleGenerate = async () => {
    if (!theme && !sourceContent) {
      error('Por favor ingresa un tema o pega el contenido base de la sesión.');
      return;
    }

    setGenerating(true);
    try {
      const res = await fetch('/api/sessions/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          theme,
          sourceContent,
          area,
          grade,
          section,
          teacherName,
          duration
        })
      });

      const data = await res.json();
      if (data.success && data.sessionData) {
        const generated: SessionPlan = {
          ...data.sessionData,
          id: `ses-${Date.now()}`,
          teacherName,
          grade,
          section,
          area,
          duration,
          theme: theme || 'Sesión Curricular',
          sourceContent,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };

        // Save to DB
        const saveRes = await fetch('/api/sessions', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(generated)
        });
        const saveData = await saveRes.json();

        setActiveSession(saveData.session || generated);
        await loadSessions();
        success(`¡Sesión oficial CNEB generada con éxito con ${data.modelUsed}!`);
      } else {
        error(data.error || 'Error al generar la sesión pedagógica.');
      }
    } catch (err: any) {
      error(err.message || 'Error en la llamada de IA.');
    } finally {
      setGenerating(false);
    }
  };

  const handleDeleteSession = async (id: string) => {
    try {
      const res = await fetch(`/api/sessions?id=${id}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        success('Sesión eliminada');
        if (activeSession?.id === id) {
          setActiveSession(null);
        }
        loadSessions();
      }
    } catch (e) {
      error('Error al eliminar');
    }
  };

  // Export to Word (.docx) con formato estructurado de tablas y cajas tipo MINEDU
  const exportToWord = async () => {
    if (!activeSession) return;

    try {
      const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' };
      const tableBorders = {
        top: cellBorder,
        bottom: cellBorder,
        left: cellBorder,
        right: cellBorder,
        insideHorizontal: cellBorder,
        insideVertical: cellBorder
      };

      const sectionHeaderRow = (title: string, rightText?: string) => new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: title, bold: true, size: 20, color: '1E293B' }),
                  ...(rightText ? [new TextRun({ text: `    (${rightText})`, italics: true, size: 18, color: '64748B' })] : [])
                ]
              })
            ],
            columnSpan: 2,
            shading: { fill: 'F1F5F9' },
            borders: tableBorders
          })
        ]
      });

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: { top: 720, bottom: 720, left: 1000, right: 1000 }
              }
            },
            children: [
              new Paragraph({
                text: "PLANIFICACIÓN DE SESIÓN DE APRENDIZAJE (MINEDU - CNEB)",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 300 },
                children: [
                  new TextRun({ text: activeSession.title, bold: true, size: 24, color: '1E3A8A' }),
                ]
              }),

              // TABLA DE DATOS INFORMATIVOS
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: tableBorders,
                rows: [
                  sectionHeaderRow("DATOS INFORMATIVOS"),
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Área Curricular: ", bold: true }), new TextRun(activeSession.area || "—")] }),
                          new Paragraph({ children: [new TextRun({ text: "Grado y Sección: ", bold: true }), new TextRun(`${activeSession.grade} "${activeSession.section}"`)] })
                        ]
                      }),
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Docente: ", bold: true }), new TextRun(activeSession.teacherName || "—")] }),
                          new Paragraph({ children: [new TextRun({ text: "Duración: ", bold: true }), new TextRun(activeSession.duration || "90 minutos")] })
                        ]
                      })
                    ]
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 2,
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Tema / Propósito General: ", bold: true }), new TextRun(activeSession.theme || "—")] })
                        ]
                      })
                    ]
                  })
                ]
              }),

              new Paragraph({ text: "", spacing: { after: 200 } }),

              // TABLA I. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: tableBorders,
                rows: [
                  sectionHeaderRow("I. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE", `Duración: ${activeSession.duration || "90 min"}`),
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Competencias", bold: true, color: '1E293B' })], spacing: { after: 80 } }),
                          ...(activeSession.competencies?.map(c => new Paragraph({ children: [new TextRun(`• ${c}`)], spacing: { after: 40 } })) || [new Paragraph({ text: "—" })])
                        ]
                      }),
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Capacidades", bold: true, color: '1E293B' })], spacing: { after: 80 } }),
                          ...(activeSession.capacities?.map(c => new Paragraph({ children: [new TextRun(`• ${c}`)], spacing: { after: 40 } })) || [new Paragraph({ text: "—" })])
                        ]
                      })
                    ]
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 2,
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Desempeños Precisados (Criterios de Evaluación)", bold: true, color: '1E293B' })], spacing: { after: 80 } }),
                          ...(activeSession.performances?.map(p => new Paragraph({ children: [new TextRun(`• ${p}`)], spacing: { after: 40 } })) || [new Paragraph({ text: "—" })])
                        ]
                      })
                    ]
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        shading: { fill: 'ECFDF5' }, // emerald tint
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Evidencia de Aprendizaje", bold: true, color: '065F46' })], spacing: { after: 60 } }),
                          new Paragraph({ children: [new TextRun({ text: activeSession.evidence || "—", color: '047857' })] })
                        ]
                      }),
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        shading: { fill: 'EEF2FF' }, // indigo tint
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Instrumento de Evaluación", bold: true, color: '3730A3' })], spacing: { after: 60 } }),
                          new Paragraph({ children: [new TextRun({ text: activeSession.instrument || "Rúbrica de evaluación analítica", color: '4338CA' })] })
                        ]
                      })
                    ]
                  })
                ]
              }),

              new Paragraph({ text: "", spacing: { after: 200 } }),

              // TABLA II. COMPETENCIAS Y ENFOQUES TRANSVERSALES
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: tableBorders,
                rows: [
                  sectionHeaderRow("II. COMPETENCIAS Y ENFOQUES TRANSVERSALES"),
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Competencias y Capacidades Transversales", bold: true, color: '1E293B' })], spacing: { after: 80 } }),
                          ...(activeSession.transversalCompetencies?.map(tc => new Paragraph({ children: [new TextRun(`• ${tc}`)], spacing: { after: 40 } })) || []),
                          ...(activeSession.transversalCapacities?.map(cap => new Paragraph({ children: [new TextRun({ text: `  - ${cap}`, italics: true })], spacing: { after: 40 } })) || [])
                        ]
                      }),
                      new TableCell({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        borders: tableBorders,
                        children: [
                          new Paragraph({ children: [new TextRun({ text: "Enfoques Transversales y Valores", bold: true, color: '1E293B' })], spacing: { after: 80 } }),
                          ...(activeSession.transversalApproaches?.map(ta => new Paragraph({ children: [new TextRun(`• ${ta}`)], spacing: { after: 40 } })) || [new Paragraph({ text: "—" })])
                        ]
                      })
                    ]
                  })
                ]
              }),

              new Paragraph({ text: "", spacing: { after: 200 } }),

              // TABLA III. SECUENCIA DIDÁCTICA (MOMENTOS PEDAGÓGICOS)
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: tableBorders,
                rows: [
                  sectionHeaderRow("III. SECUENCIA DIDÁCTICA (MOMENTOS PEDAGÓGICOS)"),
                  // Momento Inicio
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 2,
                        borders: tableBorders,
                        shading: { fill: 'FFFBEB' }, // amber tint
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: "1. MOMENTO DE INICIO ", bold: true, color: '78350F' }),
                              new TextRun({ text: `(${activeSession.sequence?.inicio?.timeMinutes || 15} minutos)`, italics: true, color: '92400E' })
                            ],
                            spacing: { after: 80 }
                          }),
                          new Paragraph({ children: [new TextRun({ text: "• Motivación: ", bold: true }), new TextRun(activeSession.sequence?.inicio?.motivation || "—")], spacing: { after: 40 } }),
                          new Paragraph({ children: [new TextRun({ text: "• Saberes Previos: ", bold: true }), new TextRun(activeSession.sequence?.inicio?.priorKnowledge || "—")], spacing: { after: 40 } }),
                          new Paragraph({ children: [new TextRun({ text: "• Conflicto Cognitivo: ", bold: true }), new TextRun(activeSession.sequence?.inicio?.cognitiveConflict || "—")], spacing: { after: 40 } }),
                          new Paragraph({ children: [new TextRun({ text: "• Propósito Comunicado: ", bold: true }), new TextRun(activeSession.sequence?.inicio?.purpose || "—")], spacing: { after: 40 } })
                        ]
                      })
                    ]
                  }),
                  // Momento Desarrollo
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 2,
                        borders: tableBorders,
                        shading: { fill: 'EFF6FF' }, // blue tint
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: "2. MOMENTO DE DESARROLLO (Gestión y Acompañamiento) ", bold: true, color: '1E3A8A' }),
                              new TextRun({ text: `(${activeSession.sequence?.desarrollo?.timeMinutes || 60} minutos)`, italics: true, color: '1D4ED8' })
                            ],
                            spacing: { after: 80 }
                          }),
                          new Paragraph({ children: [new TextRun({ text: "• Metodología Activa: ", bold: true }), new TextRun(activeSession.sequence?.desarrollo?.methodology || "—")], spacing: { after: 60 } }),
                          new Paragraph({ children: [new TextRun({ text: "Actividades de Aprendizaje:", bold: true })], spacing: { after: 40 } }),
                          ...(activeSession.sequence?.desarrollo?.activities?.map((act, i) =>
                            new Paragraph({ children: [new TextRun(`${i + 1}. ${act}`)], spacing: { after: 40 } })
                          ) || [])
                        ]
                      })
                    ]
                  }),
                  // Momento Cierre
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 2,
                        borders: tableBorders,
                        shading: { fill: 'ECFDF5' }, // emerald tint
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: "3. MOMENTO DE CIERRE ", bold: true, color: '065F46' }),
                              new TextRun({ text: `(${activeSession.sequence?.cierre?.timeMinutes || 15} minutos)`, italics: true, color: '047857' })
                            ],
                            spacing: { after: 80 }
                          }),
                          new Paragraph({ children: [new TextRun({ text: "• Metacognición: ", bold: true }), new TextRun(activeSession.sequence?.cierre?.metacognition || "—")], spacing: { after: 40 } }),
                          new Paragraph({ children: [new TextRun({ text: "• Retroalimentación y Síntesis: ", bold: true }), new TextRun(activeSession.sequence?.cierre?.evaluationSummary || "—")], spacing: { after: 40 } })
                        ]
                      })
                    ]
                  })
                ]
              }),

              new Paragraph({ text: "", spacing: { after: 200 } }),

              // TABLA IV. RÚBRICA DE EVALUACIÓN ANALÍTICA
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: tableBorders,
                rows: [
                  sectionHeaderRow("IV. RÚBRICA DE EVALUACIÓN ANALÍTICA"),
                  new TableRow({
                    children: [
                      new TableCell({ width: { size: 24, type: WidthType.PERCENTAGE }, shading: { fill: 'F8FAFC' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: "Criterio", bold: true })] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, shading: { fill: 'ECFDF5' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: "Destacado (AD)", bold: true, color: '065F46' })] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, shading: { fill: 'EFF6FF' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: "Logrado (A)", bold: true, color: '1E3A8A' })] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, shading: { fill: 'FFFBEB' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: "En Proceso (B)", bold: true, color: '92400E' })] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, shading: { fill: 'FFF1F2' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: "En Inicio (C)", bold: true, color: '9F1239' })] })] }),
                    ]
                  }),
                  ...(activeSession.rubric || []).map(r => new TableRow({
                    children: [
                      new TableCell({ width: { size: 24, type: WidthType.PERCENTAGE }, shading: { fill: 'F8FAFC' }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun({ text: r.criteria || '', bold: true })] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun(r.levels?.destacado || '—')] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun(r.levels?.logrado || '—')] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun(r.levels?.enProceso || '—')] })] }),
                      new TableCell({ width: { size: 19, type: WidthType.PERCENTAGE }, borders: tableBorders, children: [new Paragraph({ children: [new TextRun(r.levels?.enInicio || '—')] })] }),
                    ]
                  }))
                ]
              }),

              ...(activeSession.assessment?.questions && activeSession.assessment.questions.length > 0 ? [
                new Paragraph({ text: "", spacing: { after: 200 } }),
                new Table({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  borders: tableBorders,
                  rows: [
                    sectionHeaderRow("V. EVALUACIÓN Y PRUEBA OBJETIVA"),
                    new TableRow({
                      children: [
                        new TableCell({
                          columnSpan: 2,
                          borders: tableBorders,
                          children: [
                            new Paragraph({
                              children: [
                                new TextRun({ text: "Instrucciones: ", bold: true }),
                                new TextRun(activeSession.assessment.instructions || "Responde las siguientes preguntas según lo desarrollado en la sesión.")
                              ],
                              spacing: { after: 120 }
                            }),
                            ...activeSession.assessment.questions.flatMap((q, idx) => [
                              new Paragraph({
                                children: [
                                  new TextRun({ text: `${idx + 1}. ${q.question} `, bold: true }),
                                  new TextRun({ text: `(${q.points || 5} pts)`, italics: true, color: '64748B' })
                                ],
                                spacing: { before: 80, after: 40 }
                              }),
                              ...(q.options ? q.options.map(opt => new Paragraph({ children: [new TextRun(`   ${opt}`)], spacing: { after: 20 } })) : []),
                              new Paragraph({
                                children: [
                                  new TextRun({ text: "   Respuesta Correcta: ", bold: true, color: '047857' }),
                                  new TextRun({ text: q.correctAnswer || "—", color: '047857' })
                                ],
                                spacing: { after: 20 }
                              }),
                              new Paragraph({
                                children: [
                                  new TextRun({ text: "   Justificación Pedagógica: ", italics: true, color: '475569' }),
                                  new TextRun({ text: q.explanation || "—", color: '475569' })
                                ],
                                spacing: { after: 80 }
                              })
                            ])
                          ]
                        })
                      ]
                    })
                  ]
                })
              ] : [])
            ]
          }
        ]
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Sesion_CNEB_${activeSession.title.replace(/[^a-zA-Z0-9]/g, '_')}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      success('Sesión descargada en formato Word (.docx)');
    } catch (err: any) {
      error('Error al generar Word: ' + err.message);
    }
  };

  // Export to PDF con estructura idéntica visual a la aplicación (Tablas CNEB estructuradas)
  const exportToPdf = () => {
    if (!activeSession) return;
    try {
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const margin = 14;
      const contentWidth = pageWidth - (margin * 2);

      // Encabezado Principal Oficial
      pdf.setFillColor(30, 41, 59);
      pdf.setFontSize(14);
      pdf.setFont('helvetica', 'bold');
      pdf.setTextColor(30, 58, 138);
      pdf.text('PLANIFICACIÓN DE SESIÓN DE APRENDIZAJE (MINEDU - CNEB)', pageWidth / 2, 16, { align: 'center' });

      pdf.setFontSize(11);
      pdf.setTextColor(15, 23, 42);
      const titleLines = pdf.splitTextToSize(activeSession.title, contentWidth);
      pdf.text(titleLines, pageWidth / 2, 23, { align: 'center' });

      // TABLA 0: DATOS INFORMATIVOS
      autoTable(pdf, {
        startY: 28,
        margin: { left: margin, right: margin },
        head: [['DATOS INFORMATIVOS', '']],
        body: [
          [
            `Área Curricular: ${activeSession.area || '—'}\nGrado y Sección: ${activeSession.grade} "${activeSession.section}"`,
            `Docente: ${activeSession.teacherName || '—'}\nDuración: ${activeSession.duration || '90 minutos'}`
          ],
          [
            { content: `Tema / Propósito: ${activeSession.theme || '—'}`, colSpan: 2 }
          ]
        ],
        theme: 'grid',
        headStyles: {
          fillColor: [241, 245, 249],
          textColor: [30, 41, 59],
          fontStyle: 'bold',
          fontSize: 9.5
        },
        bodyStyles: {
          fontSize: 8.5,
          textColor: [51, 65, 85],
          cellPadding: 3,
          lineColor: [226, 232, 240],
          lineWidth: 0.2
        }
      });

      // TABLA I: PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE
      const compText = (activeSession.competencies || []).map(c => `• ${c}`).join('\n') || '—';
      const capText = (activeSession.capacities || []).map(c => `• ${c}`).join('\n') || '—';
      const perfText = (activeSession.performances || []).map(p => `• ${p}`).join('\n') || '—';

      autoTable(pdf, {
        startY: (pdf as any).lastAutoTable.finalY + 4,
        margin: { left: margin, right: margin },
        head: [['I. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE', `Duración: ${activeSession.duration || '90 min'}`]],
        body: [
          [
            { content: `Competencias\n${compText}`, styles: { fontStyle: 'bold', textColor: [30, 41, 59] } },
            { content: `Capacidades\n${capText}`, styles: { fontStyle: 'bold', textColor: [30, 41, 59] } }
          ],
          [
            {
              content: `Desempeños Precisados (Criterios de Evaluación)\n${perfText}`,
              colSpan: 2,
              styles: { fontStyle: 'normal', textColor: [51, 65, 85] }
            }
          ],
          [
            {
              content: `Evidencia de Aprendizaje\n${activeSession.evidence || '—'}`,
              styles: { fillColor: [236, 253, 245], textColor: [6, 95, 70] }
            },
            {
              content: `Instrumento de Evaluación\n${activeSession.instrument || 'Rúbrica de evaluación analítica'}`,
              styles: { fillColor: [238, 242, 255], textColor: [55, 48, 163] }
            }
          ]
        ],
        theme: 'grid',
        headStyles: {
          fillColor: [241, 245, 249],
          textColor: [30, 41, 59],
          fontStyle: 'bold',
          fontSize: 9.5
        },
        bodyStyles: {
          fontSize: 8.5,
          cellPadding: 3,
          lineColor: [226, 232, 240],
          lineWidth: 0.2
        }
      });

      // TABLA II: COMPETENCIAS Y ENFOQUES TRANSVERSALES
      const transCompText = [
        ...(activeSession.transversalCompetencies || []).map(tc => `• ${tc}`),
        ...(activeSession.transversalCapacities || []).map(cap => `  - ${cap}`)
      ].join('\n') || '—';

      const transAppText = (activeSession.transversalApproaches || []).map(ta => `• ${ta}`).join('\n') || '—';

      autoTable(pdf, {
        startY: (pdf as any).lastAutoTable.finalY + 4,
        margin: { left: margin, right: margin },
        head: [['II. COMPETENCIAS Y ENFOQUES TRANSVERSALES', '']],
        body: [
          [
            { content: `Competencias y Capacidades Transversales\n${transCompText}` },
            { content: `Enfoques Transversales y Valores\n${transAppText}` }
          ]
        ],
        theme: 'grid',
        headStyles: {
          fillColor: [241, 245, 249],
          textColor: [30, 41, 59],
          fontStyle: 'bold',
          fontSize: 9.5
        },
        bodyStyles: {
          fontSize: 8.5,
          textColor: [51, 65, 85],
          cellPadding: 3,
          lineColor: [226, 232, 240],
          lineWidth: 0.2
        }
      });

      // TABLA III: SECUENCIA DIDÁCTICA (MOMENTOS PEDAGÓGICOS)
      const devActivities = (activeSession.sequence?.desarrollo?.activities || [])
        .map((act, i) => `${i + 1}. ${act}`)
        .join('\n');

      autoTable(pdf, {
        startY: (pdf as any).lastAutoTable.finalY + 4,
        margin: { left: margin, right: margin },
        head: [['III. SECUENCIA DIDÁCTICA (MOMENTOS PEDAGÓGICOS)', '']],
        body: [
          [
            {
              content: `1. MOMENTO DE INICIO (${activeSession.sequence?.inicio?.timeMinutes || 15} minutos)\n• Motivación: ${activeSession.sequence?.inicio?.motivation || '—'}\n• Saberes Previos: ${activeSession.sequence?.inicio?.priorKnowledge || '—'}\n• Conflicto Cognitivo: ${activeSession.sequence?.inicio?.cognitiveConflict || '—'}\n• Propósito Comunicado: ${activeSession.sequence?.inicio?.purpose || '—'}`,
              colSpan: 2,
              styles: { fillColor: [255, 251, 235], textColor: [120, 53, 15] }
            }
          ],
          [
            {
              content: `2. MOMENTO DE DESARROLLO - Gestión y Acompañamiento (${activeSession.sequence?.desarrollo?.timeMinutes || 60} minutos)\n• Metodología Activa: ${activeSession.sequence?.desarrollo?.methodology || '—'}\n\nActividades de Aprendizaje:\n${devActivities || '—'}`,
              colSpan: 2,
              styles: { fillColor: [239, 246, 255], textColor: [30, 58, 138] }
            }
          ],
          [
            {
              content: `3. MOMENTO DE CIERRE (${activeSession.sequence?.cierre?.timeMinutes || 15} minutos)\n• Metacognición: ${activeSession.sequence?.cierre?.metacognition || '—'}\n• Retroalimentación y Síntesis: ${activeSession.sequence?.cierre?.evaluationSummary || '—'}`,
              colSpan: 2,
              styles: { fillColor: [236, 253, 245], textColor: [6, 95, 70] }
            }
          ]
        ],
        theme: 'grid',
        headStyles: {
          fillColor: [241, 245, 249],
          textColor: [30, 41, 59],
          fontStyle: 'bold',
          fontSize: 9.5
        },
        bodyStyles: {
          fontSize: 8.5,
          cellPadding: 3.5,
          lineColor: [226, 232, 240],
          lineWidth: 0.2
        }
      });

      // TABLA IV: RÚBRICA DE EVALUACIÓN ANALÍTICA
      const rubricBody = (activeSession.rubric || []).map(r => [
        { content: r.criteria || '', styles: { fontStyle: 'bold' as const, fillColor: [248, 250, 252] as [number, number, number] } },
        r.levels?.destacado || '—',
        r.levels?.logrado || '—',
        r.levels?.enProceso || '—',
        r.levels?.enInicio || '—'
      ]);

      autoTable(pdf, {
        startY: (pdf as any).lastAutoTable.finalY + 4,
        margin: { left: margin, right: margin },
        head: [
          [{ content: 'IV. RÚBRICA DE EVALUACIÓN ANALÍTICA', colSpan: 5 }],
          ['Criterio', 'Destacado (AD)', 'Logrado (A)', 'En Proceso (B)', 'En Inicio (C)']
        ],
        body: rubricBody,
        theme: 'grid',
        headStyles: {
          fillColor: [241, 245, 249],
          textColor: [30, 41, 59],
          fontStyle: 'bold',
          fontSize: 9
        },
        bodyStyles: {
          fontSize: 8,
          textColor: [51, 65, 85],
          cellPadding: 2.5,
          lineColor: [226, 232, 240],
          lineWidth: 0.2
        },
        columnStyles: {
          0: { cellWidth: 38 },
          1: { cellWidth: 36, fillColor: [240, 253, 244] },
          2: { cellWidth: 36, fillColor: [239, 246, 255] },
          3: { cellWidth: 36, fillColor: [255, 251, 235] },
          4: { cellWidth: 36, fillColor: [255, 241, 242] }
        }
      });

      // TABLA V: EVALUACIÓN Y PRUEBA OBJETIVA (Si existe)
      if (activeSession.assessment?.questions && activeSession.assessment.questions.length > 0) {
        const assessmentBody = activeSession.assessment.questions.map((q, idx) => {
          const opts = q.options ? `\n${q.options.join('\n')}` : '';
          return [
            {
              content: `${idx + 1}. ${q.question} (${q.points || 5} pts)${opts}\n• Respuesta Correcta: ${q.correctAnswer}\n• Justificación: ${q.explanation}`
            }
          ];
        });

        autoTable(pdf, {
          startY: (pdf as any).lastAutoTable.finalY + 4,
          margin: { left: margin, right: margin },
          head: [
            ['V. EVALUACIÓN Y PRUEBA OBJETIVA'],
            [`Instrucciones: ${activeSession.assessment.instructions || 'Responde las preguntas de acuerdo con lo trabajado.'}`]
          ],
          body: assessmentBody,
          theme: 'grid',
          headStyles: {
            fillColor: [241, 245, 249],
            textColor: [30, 41, 59],
            fontStyle: 'bold',
            fontSize: 9
          },
          bodyStyles: {
            fontSize: 8.5,
            textColor: [51, 65, 85],
            cellPadding: 3,
            lineColor: [226, 232, 240],
            lineWidth: 0.2
          }
        });
      }

      pdf.save(`Sesion_CNEB_${activeSession.title.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`);
      success('Sesión descargada en formato PDF');
    } catch (e: any) {
      error('Error al exportar PDF: ' + e.message);
    }
  };

  // Student test submission
  const handleAnswerChange = (qId: string, val: string) => {
    setStudentAnswers(prev => ({ ...prev, [qId]: val }));
  };

  const handleEvaluateStudentTest = () => {
    if (!activeSession?.assessment?.questions) return;
    let score = 0;
    let max = 0;
    activeSession.assessment.questions.forEach(q => {
      max += q.points || 5;
      const userAns = (studentAnswers[q.id] || '').trim().toLowerCase();
      const correctAns = (q.correctAnswer || '').trim().toLowerCase();
      if (userAns && (userAns === correctAns || correctAns.includes(userAns) || userAns.includes(correctAns))) {
        score += q.points || 5;
      }
    });
    setTestScore({ total: score, max });
    success(`Evaluación calificada: Obtuviste ${score} de ${max} puntos.`);
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
      {/* Sidebar Controls */}
      <div className="lg:col-span-4 space-y-6">
        <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
            <div className="p-2.5 bg-indigo-50 border border-indigo-100 rounded-xl text-indigo-600">
              <Sparkles className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-800">Generador CNEB con IA</h3>
              <p className="text-xs text-slate-500">Sesión oficial + Rúbrica + Examen interactivo</p>
            </div>
          </div>

          <div className="space-y-3 text-xs">
            <div>
              <label className="block font-semibold text-slate-700 mb-1">Área Curricular Oficial</label>
              <select
                value={area}
                onChange={(e) => setArea(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
              >
                <option value="Desarrollo Personal, Ciudadanía y Cívica">Desarrollo Personal, Ciudadanía y Cívica</option>
                <option value="Ciencias Sociales">Ciencias Sociales</option>
                <option value="Comunicación">Comunicación</option>
                <option value="Matemática">Matemática</option>
                <option value="Ciencia y Tecnología">Ciencia y Tecnología</option>
                <option value="Educación para el Trabajo">Educación para el Trabajo</option>
                <option value="Inglés como Lengua Extranjera">Inglés como Lengua Extranjera</option>
                <option value="Arte y Cultura">Arte y Cultura</option>
                <option value="Educación Física">Educación Física</option>
                <option value="Educación Religiosa">Educación Religiosa</option>
                <option value="Tutoría y Orientación Educativa">Tutoría y Orientación Educativa</option>
              </select>
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Grado / Nivel</label>
                <input
                  type="text"
                  value={grade}
                  onChange={(e) => setGrade(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Sección</label>
                <input
                  type="text"
                  value={section}
                  onChange={(e) => setSection(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
                />
              </div>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">Nombre del Docente</label>
              <input
                type="text"
                value={teacherName}
                onChange={(e) => setTeacherName(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">Tema Propuesto</label>
              <input
                type="text"
                placeholder="Ej. La deliberación sobre asuntos públicos y los derechos ciudadanos"
                value={theme}
                onChange={(e) => setTheme(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">O Pegar Contenido / Texto de Apoyo</label>
              <textarea
                rows={4}
                placeholder="Pega aquí párrafos de lectura, artículos, fragmentos de libros o guías para elaborar la sesión a partir de este material..."
                value={sourceContent}
                onChange={(e) => setSourceContent(e.target.value)}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
              />
            </div>

            <button
              onClick={handleGenerate}
              disabled={generating}
              className="w-full py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition-all disabled:opacity-50"
            >
              <Sparkles className={`w-4 h-4 ${generating ? 'animate-spin' : ''}`} />
              <span>{generating ? 'Generando Sesión Curricular con IA...' : 'Generar Sesión Oficial'}</span>
            </button>
          </div>
        </div>

        {/* Saved Sessions List */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-xs space-y-3">
          <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider">
            Sesiones Guardadas ({sessions.length})
          </h4>
          <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
            {sessions.length === 0 ? (
              <p className="text-xs text-slate-400 py-4 text-center">No hay sesiones creadas todavía.</p>
            ) : (
              sessions.map((s) => (
                <div
                  key={s.id}
                  onClick={() => setActiveSession(s)}
                  className={`p-3 rounded-xl border text-xs cursor-pointer transition-all flex items-center justify-between ${
                    activeSession?.id === s.id
                      ? 'bg-indigo-50/70 border-indigo-300 text-indigo-950 font-medium'
                      : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                  }`}
                >
                  <div className="truncate mr-2">
                    <p className="font-bold truncate">{s.title}</p>
                    <p className="text-[10px] text-slate-500">{s.area} • {s.grade}</p>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDeleteSession(s.id);
                    }}
                    className="text-slate-400 hover:text-rose-600 p-1 rounded transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* Main Preview and Editor Area */}
      <div className="lg:col-span-8 space-y-6">
        {activeSession ? (
          <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-6">
            {/* Header Actions */}
            <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <span className="text-[10px] uppercase font-bold tracking-wider px-2.5 py-1 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-lg">
                  Formato Oficial CNEB
                </span>
                <h2 className="text-xl font-bold text-slate-900 mt-2">{activeSession.title}</h2>
                <p className="text-xs text-slate-500">
                  {activeSession.area} • {activeSession.grade} &quot;{activeSession.section}&quot; • Docente: {activeSession.teacherName}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => {
                    setStudentAnswers({});
                    setTestScore(null);
                    setIsTestModalOpen(true);
                  }}
                  className="px-3.5 py-2 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors"
                >
                  <ListChecks className="w-4 h-4 text-amber-600" />
                  <span>Resolver Evaluación</span>
                </button>

                <button
                  onClick={exportToWord}
                  className="px-3.5 py-2 bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors"
                >
                  <FileText className="w-4 h-4 text-blue-600" />
                  <span>Descargar Word</span>
                </button>

                <button
                  onClick={exportToPdf}
                  className="px-3.5 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors"
                >
                  <Download className="w-4 h-4 text-indigo-600" />
                  <span>Descargar PDF</span>
                </button>
              </div>
            </div>

            {/* Official MINEDU/CNEB Structured Table */}
            <div className="border border-slate-200 rounded-xl overflow-hidden text-xs">
              <div className="bg-slate-100/80 px-4 py-2.5 font-bold text-slate-800 border-b border-slate-200 flex items-center justify-between">
                <span>I. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE</span>
                <span className="text-[11px] font-normal text-slate-500">Duración: {activeSession.duration}</span>
              </div>
              <div className="p-4 space-y-4 bg-white">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-200/70">
                    <p className="font-bold text-slate-800 mb-1">Competencias</p>
                    <ul className="list-disc list-inside text-slate-600 space-y-1">
                      {activeSession.competencies?.map((c, i) => <li key={i}>{c}</li>)}
                    </ul>
                  </div>
                  <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-200/70">
                    <p className="font-bold text-slate-800 mb-1">Capacidades</p>
                    <ul className="list-disc list-inside text-slate-600 space-y-1">
                      {activeSession.capacities?.map((c, i) => <li key={i}>{c}</li>)}
                    </ul>
                  </div>
                </div>

                <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-200/70">
                  <p className="font-bold text-slate-800 mb-1">Desempeños Precisados (Criterios de Evaluación)</p>
                  <ul className="list-disc list-inside text-slate-600 space-y-1">
                    {activeSession.performances?.map((p, i) => <li key={i}>{p}</li>)}
                  </ul>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="bg-emerald-50/50 p-3 rounded-xl border border-emerald-100">
                    <p className="font-bold text-emerald-900 mb-1">Evidencia de Aprendizaje</p>
                    <p className="text-emerald-800">{activeSession.evidence}</p>
                  </div>
                  <div className="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100">
                    <p className="font-bold text-indigo-900 mb-1">Instrumento de Evaluación</p>
                    <p className="text-indigo-800">{activeSession.instrument}</p>
                  </div>
                </div>
              </div>

              {/* Transversal Competencies & Approaches */}
              <div className="bg-slate-100/80 px-4 py-2.5 font-bold text-slate-800 border-t border-b border-slate-200">
                II. COMPETENCIAS Y ENFOQUES TRANSVERSALES
              </div>
              <div className="p-4 grid grid-cols-1 md:grid-cols-2 gap-4 bg-white">
                <div>
                  <p className="font-bold text-slate-800 mb-1">Competencias y Capacidades Transversales</p>
                  <ul className="list-disc list-inside text-slate-600 space-y-1">
                    {activeSession.transversalCompetencies?.map((tc, i) => <li key={i}>{tc}</li>)}
                    {activeSession.transversalCapacities?.map((tc, i) => <li key={`c-${i}`} className="italic">{tc}</li>)}
                  </ul>
                </div>
                <div>
                  <p className="font-bold text-slate-800 mb-1">Enfoques Transversales y Valores</p>
                  <ul className="list-disc list-inside text-slate-600 space-y-1">
                    {activeSession.transversalApproaches?.map((ta, i) => <li key={i}>{ta}</li>)}
                  </ul>
                </div>
              </div>

              {/* Didactic Sequence */}
              <div className="bg-slate-100/80 px-4 py-2.5 font-bold text-slate-800 border-t border-b border-slate-200">
                III. SECUENCIA DIDÁCTICA (MOMENTOS PEDAGÓGICOS)
              </div>
              <div className="p-4 space-y-4 bg-white">
                {/* Inicio */}
                <div className="p-3 bg-amber-50/40 border border-amber-200/60 rounded-xl space-y-2">
                  <div className="flex items-center justify-between font-bold text-amber-900">
                    <span>1. MOMENTO DE INICIO</span>
                    <span className="text-[11px] bg-amber-100 px-2 py-0.5 rounded">{activeSession.sequence?.inicio?.timeMinutes || 15} minutos</span>
                  </div>
                  <p><span className="font-semibold text-slate-700">• Motivación:</span> {activeSession.sequence?.inicio?.motivation}</p>
                  <p><span className="font-semibold text-slate-700">• Saberes Previos:</span> {activeSession.sequence?.inicio?.priorKnowledge}</p>
                  <p><span className="font-semibold text-slate-700">• Conflicto Cognitivo:</span> {activeSession.sequence?.inicio?.cognitiveConflict}</p>
                  <p><span className="font-semibold text-slate-700">• Propósito Comunicado:</span> {activeSession.sequence?.inicio?.purpose}</p>
                </div>

                {/* Desarrollo */}
                <div className="p-3 bg-blue-50/40 border border-blue-200/60 rounded-xl space-y-2">
                  <div className="flex items-center justify-between font-bold text-blue-900">
                    <span>2. MOMENTO DE DESARROLLO (Gestión y Acompañamiento)</span>
                    <span className="text-[11px] bg-blue-100 px-2 py-0.5 rounded">{activeSession.sequence?.desarrollo?.timeMinutes || 60} minutos</span>
                  </div>
                  <p><span className="font-semibold text-slate-700">• Metodología Activa:</span> {activeSession.sequence?.desarrollo?.methodology}</p>
                  <div className="space-y-1 pt-1">
                    <p className="font-semibold text-slate-700">Actividades de Aprendizaje:</p>
                    <ul className="list-decimal list-inside text-slate-600 space-y-1 pl-1">
                      {activeSession.sequence?.desarrollo?.activities?.map((act, i) => (
                        <li key={i}>{act}</li>
                      ))}
                    </ul>
                  </div>
                </div>

                {/* Cierre */}
                <div className="p-3 bg-emerald-50/40 border border-emerald-200/60 rounded-xl space-y-2">
                  <div className="flex items-center justify-between font-bold text-emerald-900">
                    <span>3. MOMENTO DE CIERRE</span>
                    <span className="text-[11px] bg-emerald-100 px-2 py-0.5 rounded">{activeSession.sequence?.cierre?.timeMinutes || 15} minutos</span>
                  </div>
                  <p><span className="font-semibold text-slate-700">• Metacognición:</span> {activeSession.sequence?.cierre?.metacognition}</p>
                  <p><span className="font-semibold text-slate-700">• Retroalimentación y Síntesis:</span> {activeSession.sequence?.cierre?.evaluationSummary}</p>
                </div>
              </div>

              {/* Rubric Matrix */}
              <div className="bg-slate-100/80 px-4 py-2.5 font-bold text-slate-800 border-t border-b border-slate-200">
                IV. RÚBRICA DE EVALUACIÓN ANALÍTICA
              </div>
              <div className="p-4 overflow-x-auto bg-white">
                <table className="w-full text-left text-xs border border-slate-200">
                  <thead className="bg-slate-50 text-slate-700 font-bold border-b border-slate-200">
                    <tr>
                      <th className="p-2 border-r border-slate-200 w-1/4">Criterio</th>
                      <th className="p-2 border-r border-slate-200 bg-emerald-50/40">Destacado (AD)</th>
                      <th className="p-2 border-r border-slate-200 bg-blue-50/40">Logrado (A)</th>
                      <th className="p-2 border-r border-slate-200 bg-amber-50/40">En Proceso (B)</th>
                      <th className="p-2 bg-rose-50/40">En Inicio (C)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-slate-600">
                    {activeSession.rubric?.map((r, i) => (
                      <tr key={i} className="hover:bg-slate-50/50">
                        <td className="p-2 font-bold text-slate-800 border-r border-slate-200 bg-slate-50/30">{r.criteria}</td>
                        <td className="p-2 border-r border-slate-200">{r.levels?.destacado}</td>
                        <td className="p-2 border-r border-slate-200">{r.levels?.logrado}</td>
                        <td className="p-2 border-r border-slate-200">{r.levels?.enProceso}</td>
                        <td className="p-2">{r.levels?.enInicio}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="bg-white p-12 rounded-2xl border border-slate-200/80 shadow-xs text-center space-y-3">
            <BookOpen className="w-12 h-12 text-slate-300 mx-auto" />
            <h3 className="text-base font-bold text-slate-700">Ninguna sesión seleccionada</h3>
            <p className="text-xs text-slate-400 max-w-sm mx-auto">
              Ingresa un tema o contenido en el panel izquierdo y haz clic en &quot;Generar Sesión Oficial&quot; para crear la planificación pedagógica con IA.
            </p>
          </div>
        )}
      </div>

      {/* Interactive Assessment Modal for Students */}
      <Modal
        isOpen={isTestModalOpen}
        onClose={() => setIsTestModalOpen(false)}
        title="Evaluación Interactiva de la Sesión"
        subtitle={activeSession ? `${activeSession.title} — Resuelve el examen en línea` : ''}
        maxWidth="max-w-3xl"
      >
        {activeSession?.assessment ? (
          <div className="space-y-6 text-xs text-slate-700">
            <div className="p-3 bg-indigo-50 border border-indigo-100 rounded-xl text-indigo-900 font-medium">
              {activeSession.assessment.instructions}
            </div>

            {testScore && (
              <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-emerald-900 text-sm">¡Resultado de tu Evaluación!</h4>
                  <p className="text-emerald-700 text-xs">Puntaje obtenido: {testScore.total} de {testScore.max} puntos posibles.</p>
                </div>
                <Award className="w-8 h-8 text-emerald-600" />
              </div>
            )}

            <div className="space-y-4">
              {activeSession.assessment.questions.map((q, idx) => (
                <div key={q.id} className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-3">
                  <div className="flex items-start justify-between gap-2">
                    <p className="font-bold text-slate-900 text-xs">
                      {idx + 1}. {q.question}
                    </p>
                    <span className="text-[10px] font-semibold bg-slate-200 text-slate-600 px-2 py-0.5 rounded">
                      {q.points || 5} pts
                    </span>
                  </div>

                  {q.type === 'multiple_choice' && q.options && (
                    <div className="space-y-1.5 pl-2">
                      {q.options.map((opt, oIdx) => (
                        <label key={oIdx} className="flex items-center gap-2 cursor-pointer hover:text-indigo-600">
                          <input
                            type="radio"
                            name={`q-${q.id}`}
                            value={opt}
                            checked={studentAnswers[q.id] === opt}
                            onChange={(e) => handleAnswerChange(q.id, e.target.value)}
                            className="text-indigo-600 focus:ring-indigo-500"
                          />
                          <span>{opt}</span>
                        </label>
                      ))}
                    </div>
                  )}

                  {q.type === 'true_false' && (
                    <div className="flex items-center gap-4 pl-2">
                      {['Verdadero', 'Falso'].map((val) => (
                        <label key={val} className="flex items-center gap-2 cursor-pointer hover:text-indigo-600">
                          <input
                            type="radio"
                            name={`q-${q.id}`}
                            value={val}
                            checked={studentAnswers[q.id] === val}
                            onChange={(e) => handleAnswerChange(q.id, e.target.value)}
                            className="text-indigo-600 focus:ring-indigo-500"
                          />
                          <span>{val}</span>
                        </label>
                      ))}
                    </div>
                  )}

                  {q.type === 'short_answer' && (
                    <input
                      type="text"
                      placeholder="Escribe tu respuesta aquí..."
                      value={studentAnswers[q.id] || ''}
                      onChange={(e) => handleAnswerChange(q.id, e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500 text-xs"
                    />
                  )}

                  {testScore && (
                    <div className="pt-2 border-t border-slate-200 text-[11px] text-slate-600">
                      <p><span className="font-semibold text-emerald-700">Respuesta Correcta:</span> {q.correctAnswer}</p>
                      <p className="text-slate-500 mt-0.5"><span className="font-medium">Explicación:</span> {q.explanation}</p>
                    </div>
                  )}
                </div>
              ))}
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setIsTestModalOpen(false)}
                className="px-4 py-2 font-semibold text-slate-600 hover:bg-slate-100 rounded-xl transition-colors"
              >
                Cerrar
              </button>
              <button
                onClick={handleEvaluateStudentTest}
                className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-semibold shadow-sm transition-colors"
              >
                Calificar Respuestas
              </button>
            </div>
          </div>
        ) : (
          <p className="text-xs text-slate-400">No hay evaluación disponible para esta sesión.</p>
        )}
      </Modal>
    </div>
  );
}
