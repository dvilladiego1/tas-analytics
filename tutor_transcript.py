#!/usr/bin/env python3
"""
Profesor interactivo para el podcast:
"From Copilots to Agents: Rebuilding the Company Around AI"

Uso:
    python tutor_transcript.py

Comandos especiales:
    /quiz      — El profesor te hace una pregunta de comprensión
    /resumen   — Resumen de los temas más importantes
    /temas     — Lista los temas principales del podcast
    /reset     — Reinicia la conversación
    /salir     — Sale del programa
"""

import anthropic
import sys
from pathlib import Path

TRANSCRIPT_PATH = Path("/Users/danielvilladiego/Desktop/Brain View/From Copilots to Agents_ Rebuilding the Company Around AI.txt")

SYSTEM_PROMPT = """Eres un profesor experto en transformación digital, inteligencia artificial aplicada a negocios, y estrategia empresarial. Tu misión es enseñarle al alumno el contenido de la siguiente transcripción de podcast de forma profunda y memorable.

**Tu metodología de enseñanza:**
1. **Socrática**: Haz preguntas para guiar al alumno a descubrir los conceptos por sí mismo
2. **Contextualizada**: Conecta los conceptos con ejemplos del mundo real y con la experiencia del alumno en Kavak/TAS
3. **Progresiva**: Comienza por los conceptos fundamentales antes de avanzar a los complejos
4. **Activa**: Propón ejercicios mentales, casos de estudio y reflexiones

**Cuando el alumno pregunta algo:**
- Explica el concepto con claridad y usa una analogía si aplica
- Conecta con otros conceptos del podcast
- Pregunta al final si quiere profundizar o tiene dudas

**Cuando el alumno pide /quiz:**
- Formula UNA pregunta de comprensión relevante sobre el contenido
- Espera la respuesta del alumno antes de evaluar

**Cuando el alumno pide /resumen:**
- Da un resumen estructurado con los 5-7 puntos clave más importantes del podcast

**Cuando el alumno pide /temas:**
- Lista todos los temas principales cubiertos en la transcripción

**Tono:** Amigable, directo, motivador. Trátalo como un colega inteligente que quiere aprender rápido.

---

**TRANSCRIPCIÓN DEL PODCAST:**

"""

def load_transcript() -> str:
    if not TRANSCRIPT_PATH.exists():
        print(f"Error: No se encontró el archivo en {TRANSCRIPT_PATH}")
        sys.exit(1)
    return TRANSCRIPT_PATH.read_text(encoding="utf-8")

def run_tutor():
    client = anthropic.Anthropic()
    transcript = load_transcript()

    # El system prompt incluye el transcript cacheado
    system_with_transcript = SYSTEM_PROMPT + transcript

    messages = []

    print("=" * 60)
    print("  PROFESOR IA — Copilots to Agents (Kavak)")
    print("=" * 60)
    print("Comandos: /quiz  /resumen  /temas  /reset  /salir")
    print("Escribe tu pregunta o tema que quieres aprender.\n")

    # Mensaje de bienvenida del profesor
    welcome_messages = [{"role": "user", "content": "Hola profesor, estoy listo para aprender. ¿Por dónde empezamos?"}]

    print("Profesor: ", end="", flush=True)
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=1024,
        system=[
            {
                "type": "text",
                "text": system_with_transcript,
                "cache_control": {"type": "ephemeral"}
            }
        ],
        messages=welcome_messages,
        thinking={"type": "adaptive"},
    ) as stream:
        full_response = ""
        for text in stream.text_stream:
            print(text, end="", flush=True)
            full_response += text

    print("\n")
    messages.append({"role": "user", "content": "Hola profesor, estoy listo para aprender. ¿Por dónde empezamos?"})
    messages.append({"role": "assistant", "content": full_response})

    # Loop principal
    while True:
        try:
            user_input = input("Tú: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\nHasta luego!")
            break

        if not user_input:
            continue

        if user_input.lower() in ("/salir", "/exit", "/quit"):
            print("Profesor: ¡Hasta la próxima! Sigue aplicando lo aprendido. 🚀")
            break

        if user_input.lower() == "/reset":
            messages = []
            print("Profesor: Conversación reiniciada. ¿Por dónde quieres empezar?\n")
            continue

        # Comandos especiales → los pasamos directamente como mensaje
        messages.append({"role": "user", "content": user_input})

        print("\nProfesor: ", end="", flush=True)

        try:
            with client.messages.stream(
                model="claude-opus-4-6",
                max_tokens=2048,
                system=[
                    {
                        "type": "text",
                        "text": system_with_transcript,
                        "cache_control": {"type": "ephemeral"}
                    }
                ],
                messages=messages,
                thinking={"type": "adaptive"},
            ) as stream:
                full_response = ""
                for text in stream.text_stream:
                    print(text, end="", flush=True)
                    full_response += text

            print("\n")
            messages.append({"role": "assistant", "content": full_response})

        except anthropic.RateLimitError:
            print("\n[Rate limit alcanzado, espera un momento...]\n")
            messages.pop()  # Quita el último mensaje del usuario
        except anthropic.APIConnectionError:
            print("\n[Error de conexión, intenta de nuevo]\n")
            messages.pop()
        except Exception as e:
            print(f"\n[Error: {e}]\n")
            messages.pop()

if __name__ == "__main__":
    run_tutor()
