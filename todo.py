import json
import os

ARCHIVO = "tareas.json"

def cargar_tareas():
    if os.path.exists(ARCHIVO):
        with open(ARCHIVO, 'r') as f:
            return json.load(f)
    return []

def guardar_tareas(tareas):
    with open(ARCHIVO, 'w') as f:
        json.dump(tareas, f, indent=4)

def mostrar_menu():
    print("\n" + "="*30)
    print("   📋 GESTOR DE TAREAS VSCode   ")
    print("="*30)
    print("1. Ver tareas")
    print("2. Añadir tarea")
    print("3. Completar tarea")
    print("4. Salir")

def main():
    tareas = cargar_tareas()
    while True:
        mostrar_menu()
        opcion = input("Elige una opción: ")

        if opcion == "1":
            print("\n--- TUS TAREAS ---")
            if not tareas:
                print("🎉 No hay tareas pendientes. ¡Genial!")
            else:
                for i, t in enumerate(tareas):
                    estado = "✅" if t['completada'] else "⏳"
                    print(f"{i+1}. {estado} {t['descripcion']}")
        
        elif opcion == "2":
            desc = input("Describe la nueva tarea: ")
            tareas.append({"descripcion": desc, "completada": False})
            guardar_tareas(tareas)
            print(f"✔️ Tarea '{desc}' añadida.")
        
        elif opcion == "3":
            if not tareas:
                print("No hay tareas que completar.")
                continue
            try:
                num = int(input("Número de tarea a marcar como hecha: ")) - 1
                if 0 <= num < len(tareas):
                    tareas[num]['completada'] = True
                    guardar_tareas(tareas)
                    print("✅ ¡Bien hecho! Tarea completada.")
                else:
                    print("Número inválido.")
            except ValueError:
                print("Por favor, introduce un número.")
        
        elif opcion == "4":
            print("👋 ¡Hasta luego! (Los datos se guardaron en 'tareas.json')")
            break
        else:
            print("Opción no válida.")

if __name__ == "__main__":
    main()