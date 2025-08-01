import flet as ft
import subprocess
import threading
import queue
import os
import time
from typing import Dict, List

class Terminal:
    def __init__(self, terminal_id: str, on_output_callback, on_directory_change_callback=None):
        self.terminal_id = terminal_id
        self.on_output_callback = on_output_callback
        self.on_directory_change_callback = on_directory_change_callback
        self.process = None
        self.output_queue = queue.Queue()
        self.input_queue = queue.Queue()
        self.is_running = False
        self.current_directory = os.getcwd()
        self.last_command = ""
        
    def start(self):
        """Inicia el proceso de terminal"""
        if self.is_running:
            return
            
        self.is_running = True
        
        # Configurar el proceso según el sistema operativo
        if os.name == 'nt':  # Windows
            self.process = subprocess.Popen(
                ['cmd'],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,
                cwd=self.current_directory
            )
        else:  # Linux/Mac
            self.process = subprocess.Popen(
                ['/bin/bash'],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=0,
                cwd=self.current_directory
            )
        
        # Iniciar hilos para manejar entrada y salida
        threading.Thread(target=self._read_output, daemon=True).start()
        threading.Thread(target=self._write_input, daemon=True).start()
        
        # Mostrar prompt inicial
        self.on_output_callback(f"Terminal {self.terminal_id} iniciada\n")
        
    def _read_output(self):
        """Lee la salida del proceso"""
        while self.is_running and self.process and self.process.poll() is None:
            try:
                output = self.process.stdout.readline()
                if output:
                    self.on_output_callback(output)
                time.sleep(0.01)
            except Exception as e:
                break
                
    def _write_input(self):
        """Escribe entrada al proceso"""
        while self.is_running and self.process:
            try:
                if not self.input_queue.empty():
                    command = self.input_queue.get()
                    if self.process and self.process.stdin:
                        self.process.stdin.write(command + '\n')
                        self.process.stdin.flush()
                time.sleep(0.01)
            except Exception as e:
                break
                
    def send_command(self, command: str):
        """Envía un comando al terminal"""
        if self.is_running:
            self.last_command = command.strip().lower()
            self.input_queue.put(command)
            
            # Detectar comandos que cambian directorio
            if self.last_command.startswith('cd ') or self.last_command == 'cd':
                # Programar actualización del directorio después de un momento
                threading.Thread(target=self._update_directory_delayed, daemon=True).start()
            
    def stop(self):
        """Detiene el terminal"""
        self.is_running = False
        if self.process:
            self.process.terminate()
            self.process = None
            
    def _update_directory_delayed(self):
        """Actualiza el directorio actual después de un comando cd"""
        time.sleep(0.5)  # Esperar a que el comando se ejecute
        try:
            if self.process and self.process.poll() is None:
                # Enviar comando para obtener directorio actual
                if os.name == 'nt':  # Windows
                    self.process.stdin.write('echo %CD%\n')
                else:  # Linux/Mac
                    self.process.stdin.write('pwd\n')
                self.process.stdin.flush()
                
                # Leer la respuesta
                time.sleep(0.2)
                if self.process.stdout:
                    response = self.process.stdout.readline().strip()
                    if response and response != self.current_directory:
                        self.current_directory = response
                        if self.on_directory_change_callback:
                            self.on_directory_change_callback(self.current_directory)
        except:
            pass
            
    def get_current_directory(self):
        """Obtiene el directorio actual"""
        return os.path.basename(self.current_directory) if self.current_directory else "~"

class TerminalTab:
    def __init__(self, terminal_id: str, on_close_callback, page):
        self.terminal_id = terminal_id
        self.on_close_callback = on_close_callback
        self.page = page
        self.current_directory = "~"
        
        # Crear el texto del título que se actualizará
        self.title_text = ft.Text(
            f"🖥 Terminal {terminal_id} ({self.current_directory})",
            size=16,
            weight=ft.FontWeight.BOLD,
            color="#00ff00"
        )
        
        self.terminal = Terminal(terminal_id, self._on_terminal_output, self._on_directory_change)
        
        # Crear scroll para la salida del terminal
        self.output_container = ft.Column(
            controls=[],
            scroll=ft.ScrollMode.ALWAYS,
            auto_scroll=True,
            spacing=0
        )
        
        # Container para el scroll
        self.scroll_container = ft.Container(
            content=self.output_container,
            bgcolor="#0c0c0c",
            border=ft.border.all(1, "#003300"),
            border_radius=5,
            padding=10,
            height=400,
            expand=True
        )
        
        self.input_field = ft.TextField(
            hint_text="Escribe aquí tu comando...",
            bgcolor="#0c0c0c",
            color="#00ff00",
            border_color="#003300",
            on_submit=self._on_command_submit,
            text_style=ft.TextStyle(
                font_family="Consolas, monospace",
                size=12
            ),
            expand=True
        )
        
        # Botones de control
        self.send_button = ft.ElevatedButton(
            text="▶",
            on_click=self._on_send_click,
            bgcolor="#003300",
            color="#00ff00",
            width=50
        )
        
        self.clear_button = ft.ElevatedButton(
            text="🗑",
            on_click=self._on_clear_click,
            bgcolor="#330000",
            color="#ff0000",
            width=50
        )
        
        self.close_button = ft.ElevatedButton(
            text="✖",
            on_click=self._on_close_click,
            bgcolor="#330000",
            color="#ff0000",
            width=50
        )
        
        # Crear el contenido de la pestaña
        self.content = ft.Container(
            content=ft.Column([
                ft.Row([
                    self.title_text,
                    ft.Row([
                        self.clear_button,
                        self.close_button
                    ])
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                
                self.scroll_container,
                
                ft.Row([
                    self.input_field,
                    self.send_button
                ], spacing=10)
            ], 
            expand=True,
            spacing=10),
            padding=20,
            bgcolor="#0c0c0c",
            border_radius=10,
            expand=True
        )
        
    def start_terminal(self):
        """Inicia el terminal después de que la UI esté lista"""
        def delayed_start():
            time.sleep(0.3)
            self.terminal.start()
            # Dar foco al campo de entrada
            time.sleep(0.1)
            if self.page:
                self.page.update()
        
        threading.Thread(target=delayed_start, daemon=True).start()
        
    def _on_directory_change(self, new_directory):
        """Callback cuando cambia el directorio"""
        # Obtener solo el nombre del directorio actual
        dir_name = os.path.basename(new_directory) if new_directory else "~"
        if not dir_name:  # Si es la raíz
            dir_name = "/"
        
        self.current_directory = dir_name
        self.title_text.value = f"🖥 Terminal {self.terminal_id} ({dir_name})"
        
        if self.page:
            try:
                self.title_text.update()
            except:
                pass
                
    def _on_terminal_output(self, output: str):
        """Callback para cuando el terminal produce salida"""
        # Agregar línea de salida
        output_line = ft.Text(
            output.rstrip(),
            color="#00ff00",
            font_family="Consolas, monospace",
            size=12,
            selectable=True
        )
        
        self.output_container.controls.append(output_line)
        
        # Mantener solo las últimas 500 líneas
        if len(self.output_container.controls) > 500:
            self.output_container.controls = self.output_container.controls[-500:]
        
        if self.page:
            try:
                self.page.update()
            except:
                pass
        
    def _on_command_submit(self, e):
        """Callback para cuando se presiona Enter en el campo de entrada"""
        self._send_command()
        
    def _on_send_click(self, e):
        """Callback para el botón enviar"""
        self._send_command()
        
    def _send_command(self):
        """Envía el comando al terminal"""
        command = self.input_field.value.strip()
        if command:
            # Mostrar comando en la salida
            command_line = ft.Text(
                f"~$ {command}",
                color="#ffff00",  # Amarillo para los comandos
                font_family="Consolas, monospace",
                size=12,
                weight=ft.FontWeight.BOLD,
                selectable=True
            )
            
            self.output_container.controls.append(command_line)
            
            # Enviar comando al terminal
            self.terminal.send_command(command)
            
            # Limpiar campo de entrada
            self.input_field.value = ""
            
            # Mantener solo las últimas 500 líneas
            if len(self.output_container.controls) > 500:
                self.output_container.controls = self.output_container.controls[-500:]
            
            if self.page:
                try:
                    self.page.update()
                except:
                    pass
            
    def _on_clear_click(self, e):
        """Limpia la salida del terminal"""
        self.output_container.controls.clear()
        if self.page:
            try:
                self.page.update()
            except:
                pass
        
    def _on_close_click(self, e):
        """Cierra esta pestaña de terminal"""
        self.terminal.stop()
        self.on_close_callback(self.terminal_id)

class TerminalManager:
    def __init__(self, page: ft.Page):
        self.page = page
        self.terminals: Dict[str, TerminalTab] = {}
        self.terminal_counter = 0
        
        # Configurar la página
        self.page.title = "🖥 Terminal Manager"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.window_width = 1200
        self.page.window_height = 800
        self.page.bgcolor = "#000000"
        
        # Crear tabs container
        self.tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[],
            expand=True,
            tab_alignment=ft.TabAlignment.START
        )
        
        # Crear barra de herramientas
        self.toolbar = ft.Row([
            ft.Text("🖥 Terminal Manager", size=20, color="#00ff00", weight=ft.FontWeight.BOLD),
            ft.Row([
                ft.ElevatedButton(
                    text="+ Nueva Terminal",
                    on_click=self._add_terminal,
                    bgcolor="#003300",
                    color="#00ff00"
                ),
                ft.ElevatedButton(
                    text="🔄 Actualizar",
                    on_click=self._refresh_all,
                    bgcolor="#003300",
                    color="#00ff00"
                )
            ])
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        
        # Configurar el layout principal
        self.main_container = ft.Container(
            content=ft.Column([
                ft.Container(
                    content=self.toolbar,
                    bgcolor="#0c0c0c",
                    padding=15,
                    border_radius=ft.border_radius.only(top_left=10, top_right=10)
                ),
                ft.Container(
                    content=self.tabs,
                    expand=True,
                    padding=10,
                    bgcolor="#000000"
                )
            ]),
            expand=True,
            margin=10,
            border_radius=10,
            bgcolor="#000000"
        )
        
        # Agregar a la página
        self.page.add(self.main_container)
        
        # Agregar primera terminal
        self._add_terminal(None)
        
    def _add_terminal(self, e):
        """Agrega una nueva terminal"""
        self.terminal_counter += 1
        terminal_id = str(self.terminal_counter)
        
        terminal_tab = TerminalTab(terminal_id, self._close_terminal, self.page)
        self.terminals[terminal_id] = terminal_tab
        
        # Crear pestaña
        tab = ft.Tab(
            text=f"Terminal {terminal_id}",
            content=terminal_tab.content
        )
        
        self.tabs.tabs.append(tab)
        self.tabs.selected_index = len(self.tabs.tabs) - 1
        self.page.update()
        
        # Iniciar terminal después de que se agregue a la página
        terminal_tab.start_terminal()
        
    def _close_terminal(self, terminal_id: str):
        """Cierra una terminal específica"""
        if terminal_id in self.terminals:
            # Encontrar y eliminar la pestaña
            for i, tab in enumerate(self.tabs.tabs):
                if f"Terminal {terminal_id}" in tab.text:
                    self.tabs.tabs.pop(i)
                    break
            
            # Eliminar el terminal del diccionario
            del self.terminals[terminal_id]
            
            # Si no quedan terminales, agregar una nueva
            if len(self.terminals) == 0:
                self._add_terminal(None)
            else:
                # Ajustar el índice seleccionado
                if self.tabs.selected_index >= len(self.tabs.tabs):
                    self.tabs.selected_index = len(self.tabs.tabs) - 1
                    
            self.page.update()
            
    def _refresh_all(self, e):
        """Actualiza todas las terminales"""
        try:
            self.page.update()
        except:
            pass

def main(page: ft.Page):
    terminal_manager = TerminalManager(page)

if __name__ == "__main__":
    ft.app(target=main)