import flet as ft
from datetime import datetime
import random
import sqlite3
import os
import hashlib
import time
import subprocess
import sys
import threading
import importlib.util
from typing import Optional, List, Dict

class DatabaseManager:
    def __init__(self):
        # Crear base de datos en el escritorio
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.db_path = os.path.join(desktop, "executive_dashboard.db")
        self.init_database()
    
    def init_database(self):
        """Inicializar la base de datos con las tablas necesarias"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Crear tabla de usuarios
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                employee_id INTEGER PRIMARY KEY,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'production')),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Insertar usuario admin por defecto si no existe
        cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
        if cursor.fetchone()[0] == 0:
            # Guardar directamente sin hash para el admin por defecto
            cursor.execute('''
                INSERT INTO users (employee_id, first_name, last_name, password, role)
                VALUES (1000, 'Admin', 'Default', 'AD99', 'admin')
            ''')
        
        conn.commit()
        conn.close()
    
    def hash_password(self, password: str) -> str:
        """Hash de la contraseña"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def generate_password(self, first_name: str, last_name: str) -> str:
        """Generar contraseña: primera letra nombre + primera letra apellido + número random"""
        random_num = random.randint(1, 99)
        password = f"{first_name[0].upper()}{last_name[0].upper()}{random_num:02d}"
        return password
    
    def authenticate_user(self, employee_id: int, password: str) -> Optional[Dict]:
        """Autenticar usuario"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Para el usuario admin por defecto, no usar hash
        if employee_id == 1000:
            cursor.execute('''
                SELECT employee_id, first_name, last_name, role 
                FROM users 
                WHERE employee_id = ? AND password = ?
            ''', (employee_id, password))
        else:
            # Para usuarios creados después, usar hash
            hashed_password = self.hash_password(password)
            cursor.execute('''
                SELECT employee_id, first_name, last_name, role 
                FROM users 
                WHERE employee_id = ? AND password = ?
            ''', (employee_id, hashed_password))
        
        user = cursor.fetchone()
        conn.close()
        
        if user:
            return {
                'employee_id': user[0],
                'first_name': user[1],
                'last_name': user[2],
                'role': user[3]
            }
        return None
    
    def create_user(self, employee_id: int, first_name: str, last_name: str, role: str) -> str:
        """Crear nuevo usuario y retornar la contraseña generada"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            password = self.generate_password(first_name, last_name)
            hashed_password = self.hash_password(password)
            
            cursor.execute('''
                INSERT INTO users (employee_id, first_name, last_name, password, role)
                VALUES (?, ?, ?, ?, ?)
            ''', (employee_id, first_name, last_name, hashed_password, role))
            
            conn.commit()
            conn.close()
            return password
        except sqlite3.IntegrityError:
            conn.close()
            raise Exception("El número de empleado ya existe")
    
    def get_all_users(self) -> List[Dict]:
        """Obtener todos los usuarios"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT employee_id, first_name, last_name, role, created_at
            FROM users
            ORDER BY employee_id
        ''')
        
        users = []
        for row in cursor.fetchall():
            users.append({
                'employee_id': row[0],
                'first_name': row[1],
                'last_name': row[2],
                'role': row[3],
                'created_at': row[4]
            })
        
        conn.close()
        return users
    
    def delete_user(self, employee_id: int) -> bool:
        """Eliminar usuario"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # No permitir eliminar el último admin
        cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
        admin_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT role FROM users WHERE employee_id = ?", (employee_id,))
        user_role = cursor.fetchone()
        
        if user_role and user_role[0] == 'admin' and admin_count <= 1:
            conn.close()
            return False
        
        cursor.execute("DELETE FROM users WHERE employee_id = ?", (employee_id,))
        conn.commit()
        conn.close()
        return True

class ExecutiveDashboard:
    def __init__(self):
        self.current_index = 0
        self.theme_color = ft.Colors.BLUE_700
        self.db = DatabaseManager()
        self.current_user = None
        self.sidebar_expanded = True
        
    def main(self, page: ft.Page):
        self.page = page
        page.title = "MASTER PLAN"
        page.theme_mode = ft.ThemeMode.DARK
        page.padding = 0
        # Ventana pequeña para el login
        page.window_width = 450
        page.window_height = 600
        page.window_center = True
        
        # Iniciar con la pantalla de login
        self.show_login_screen()
    
    def logout(self):
        """Cerrar sesión y volver al login"""
        self.current_user = None
        self.page.window_width = 450
        self.page.window_height = 600
        self.page.window_center = True
        self.page.title = "IRP--TOOLS"
        self.show_login_screen()
    
    def show_snackbar_simple(self, message, bgcolor=ft.Colors.BLUE_700):
        """Método helper para mostrar snackbars simples"""
        snack_bar = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=bgcolor
        )
        self.page.snack_bar = snack_bar
        snack_bar.open = True
        self.page.update()

    def launch_credit_memo_dashboard(self, e):
        """Lanza el dashboard de Credit Memos"""
        try:
            # Mostrar mensaje de carga
            self.show_snackbar_simple("🚀 Iniciando Credit Memo Dashboard...", ft.Colors.BLUE_700)
            
            # Buscar archivo cm.py en el directorio actual
            credit_memo_file = "cm.py"
            
            if not os.path.exists(credit_memo_file):
                self.show_snackbar_simple("❌ Archivo cm.py no encontrado", ft.Colors.RED_400)
                print(f"❌ No se encontró el archivo: {credit_memo_file}")
                print(f"📁 Archivos en directorio: {[f for f in os.listdir('.') if f.endswith('.py')]}")
                return

            def run_external_app():
                try:
                    # Ejecutar cm.py en nueva ventana
                    process = subprocess.Popen([
                        sys.executable, credit_memo_file
                    ], 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0
                    )
                    
                    print(f"✅ Credit Memo Dashboard (cm.py) iniciado con PID: {process.pid}")
                    
                    # Verificar errores inmediatos
                    time.sleep(2)
                    if process.poll() is not None:
                        stdout, stderr = process.communicate()
                        if stderr:
                            print(f"❌ Error en cm.py: {stderr.decode()}")
                        
                except Exception as e:
                    print(f"❌ Error ejecutando cm.py: {e}")
            
            # Ejecutar en hilo separado
            credit_thread = threading.Thread(target=run_external_app, daemon=True)
            credit_thread.start()
            
            # Mostrar confirmación
            self.show_snackbar_simple("✅ Credit Memo Dashboard iniciado", ft.Colors.GREEN_700)
            
        except Exception as ex:
            error_msg = f"❌ Error: {str(ex)}"
            self.show_snackbar_simple(error_msg, ft.Colors.RED_400)
            print(f"❌ Error ejecutando Credit Memo Dashboard: {ex}")

    def launch_kiteo_dashboard(self, e):
        """Lanza el dashboard de Grupos/Kiteo"""
        try:
            # Mostrar mensaje de carga
            self.show_snackbar_simple("🚀 Iniciando Grupos/Kiteo Dashboard...", ft.Colors.PURPLE_700)
            
            # Buscar archivo grupo.py en el directorio actual
            kiteo_file = "grupo.py"
            
            if not os.path.exists(kiteo_file):
                error_msg = f"❌ Archivo {kiteo_file} no encontrado"
                self.show_snackbar_simple(error_msg, ft.Colors.RED_400)
                print(error_msg)
                print(f"📁 Archivos Python en directorio: {[f for f in os.listdir('.') if f.endswith('.py')]}")
                return

            def run_external_app():
                try:
                    print(f"🚀 Ejecutando: {kiteo_file}")
                    process = subprocess.Popen([
                        sys.executable, kiteo_file
                    ], 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0
                    )
                    
                    print(f"✅ Grupos/Kiteo Dashboard ({kiteo_file}) iniciado con PID: {process.pid}")
                    
                    # Verificar errores inmediatos
                    time.sleep(3)
                    if process.poll() is not None:
                        stdout, stderr = process.communicate()
                        if stderr:
                            print(f"❌ Error en {kiteo_file}: {stderr.decode()}")
                        else:
                            print(f"✅ {kiteo_file} ejecutado correctamente")
                        
                except Exception as e:
                    print(f"❌ Error ejecutando {kiteo_file}: {e}")
            
            # Ejecutar en hilo separado
            kiteo_thread = threading.Thread(target=run_external_app, daemon=True)
            kiteo_thread.start()
            
            # Mostrar confirmación
            self.show_snackbar_simple(f"✅ Grupos/Kiteo Dashboard iniciado", ft.Colors.GREEN_700)
            
        except Exception as ex:
            error_msg = f"❌ Error: {str(ex)}"
            self.show_snackbar_simple(error_msg, ft.Colors.RED_400)
            print(f"❌ Error ejecutando Grupos/Kiteo Dashboard: {ex}")

    def launch_vew2_dashboard(self, e):
        """Botón de referencia - No funcional por ahora"""
        self.show_snackbar_simple("🔧 Función en desarrollo - Solo referencia", ft.Colors.ORANGE_400)

    def show_login_screen(self):
        """Mostrar pantalla de login"""
        self.page.clean()
        
        # Texto para mensajes
        message_text = ft.Text(
            "Ingrese sus credenciales",
            size=14,
            color=ft.Colors.WHITE54,
            text_align=ft.TextAlign.CENTER,
        )
        
        # Campos
        employee_id_field = ft.TextField(
            label="Número de Empleado",
            width=300,
            keyboard_type=ft.KeyboardType.NUMBER,
            autofocus=True,
        )
        
        password_field = ft.TextField(
            label="Contraseña",
            width=300,
            password=True,
            can_reveal_password=True,
        )
        
        def handle_login(e):
            message_text.value = "Verificando..."
            message_text.color = ft.Colors.BLUE_400
            self.page.update()
            
            try:
                if not employee_id_field.value or not password_field.value:
                    message_text.value = "⚠️ Complete todos los campos"
                    message_text.color = ft.Colors.ORANGE_400
                    self.page.update()
                    return
                
                employee_id = int(employee_id_field.value)
                password = password_field.value
                
                user = self.db.authenticate_user(employee_id, password)
                if user:
                    self.current_user = user
                    message_text.value = f"✅ ¡Bienvenido {user['first_name']}!"
                    message_text.color = ft.Colors.GREEN_400
                    self.page.update()
                    time.sleep(1)
                    self.show_dashboard()
                else:
                    message_text.value = "❌ Credenciales incorrectas"
                    message_text.color = ft.Colors.RED_400
                    self.page.update()
                    
            except ValueError:
                message_text.value = "❌ Número de empleado inválido"
                message_text.color = ft.Colors.RED_400
                self.page.update()
        
        employee_id_field.on_submit = lambda _: password_field.focus()
        password_field.on_submit = handle_login
        
        info_card = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(ft.Icons.INFO_OUTLINE, size=20, color=ft.Colors.BLUE_400),
                    ft.Text("Credenciales de Admin", size=14, weight=ft.FontWeight.W_500),
                ]),
                ft.Container(height=5),
                ft.Text("Empleado: 1000", size=12, color=ft.Colors.WHITE70),
                ft.Text("Contraseña: AD99", size=12, color=ft.Colors.WHITE70),
            ]),
            padding=15,
            bgcolor=ft.Colors.BLUE_900,
            border_radius=10,
            width=300,
        )
        
        login_container = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.BUSINESS, size=80, color=self.theme_color),
                ft.Text("Kiting Plan", size=32, weight=ft.FontWeight.BOLD),
                ft.Text("Sistema de Kiteo EZAir", size=14, color=ft.Colors.WHITE54),
                ft.Container(height=30),
                message_text,
                ft.Container(height=20),
                employee_id_field,
                password_field,
                ft.Container(height=20),
                ft.ElevatedButton(
                    "Iniciar Sesión",
                    width=300,
                    height=50,
                    bgcolor=self.theme_color,
                    color=ft.Colors.WHITE,
                    on_click=handle_login,
                ),
                ft.Container(height=30),
                info_card,
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            width=400,
            padding=40,
            bgcolor=ft.Colors.GREY_900,
            border_radius=20,
        )
        
        self.page.add(
            ft.Container(
                content=login_container,
                expand=True,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.BLACK,
            )
        )

    def show_dashboard(self):
        """Mostrar el dashboard principal"""
        # Expandir la ventana para el dashboard
        self.page.window_width = 1200
        self.page.window_height = 800
        self.page.window_center = True
        self.page.title = f"IRP-TOOL KITING - {self.current_user['first_name']} {self.current_user['last_name']}"
        
        self.page.clean()
        
        # Función para cambiar tabs con animación
        def change_tab(index):
            self.current_index = index
            # Actualizar indicador activo
            for i, indicator in enumerate(nav_indicators):
                indicator.opacity = 1 if i == index else 0
            
            # Actualizar contenido
            content_container.content = create_content(index)
            self.page.update()
        
        # Crear indicadores de navegación
        nav_indicators = []
        
        # Navegación lateral elegante con permisos
        nav_items = [
            {"icon": ft.Icons.DASHBOARD, "label": "Dashboard", "index": 0, "permission": "all"},
            {"icon": ft.Icons.ANALYTICS, "label": "Analytics", "index": 1, "permission": "all"},
            {"icon": ft.Icons.PEOPLE, "label": "Team", "index": 2, "permission": "all"},
            {"icon": ft.Icons.ATTACH_MONEY, "label": "Finance", "index": 3, "permission": "admin"},
            {"icon": ft.Icons.SETTINGS, "label": "Settings", "index": 4, "permission": "all"},
        ]
        
        # Agregar opciones de admin si el usuario es admin
        if self.current_user['role'] == 'admin':
            nav_items.extend([
                {"icon": ft.Icons.ADMIN_PANEL_SETTINGS, "label": "Admin", "index": 5, "permission": "admin"},
                {"icon": ft.Icons.MANAGE_ACCOUNTS, "label": "Usuarios", "index": 6, "permission": "admin"}
            ])
        
        def create_nav_item(icon, label, index, permission):
            # Verificar permisos
            if permission == "admin" and self.current_user['role'] != 'admin':
                return None
            
            indicator = ft.Container(
                width=4,
                height=40,
                bgcolor=self.theme_color,
                border_radius=ft.border_radius.only(top_right=10, bottom_right=10),
                opacity=1 if index == 0 else 0,
                animate_opacity=300,
            )
            nav_indicators.append(indicator)
            
            nav_icon = ft.Icon(icon, color=ft.Colors.WHITE70, size=24)
            nav_text = ft.Text(label, size=16, weight=ft.FontWeight.W_400, color=ft.Colors.WHITE70)
            
            return ft.Container(
                content=ft.Row([
                    indicator,
                    ft.Container(width=20),
                    nav_icon,
                    ft.Container(width=15),
                    ft.Container(
                        content=nav_text,
                        width=0 if not self.sidebar_expanded else 150,
                        animate=ft.Animation(200, ft.AnimationCurve.EASE_IN_OUT),
                    ),
                ]),
                height=60,
                on_click=lambda _: change_tab(index),
                animate=ft.Animation(200, ft.AnimationCurve.EASE_IN_OUT),
            )
        
        # Función para expandir/contraer sidebar
        def toggle_sidebar(e):
            self.sidebar_expanded = e.data == "true"
            sidebar.width = 250 if self.sidebar_expanded else 80
            
            # Actualizar visibilidad de textos
            for item in nav_container.controls:
                if item and hasattr(item, 'content') and hasattr(item.content, 'controls'):
                    if len(item.content.controls) > 4:
                        text_container = item.content.controls[4]
                        text_container.width = 150 if self.sidebar_expanded else 0
            
            # Actualizar logo/header
            logo_text.visible = self.sidebar_expanded
            subtitle_text.visible = self.sidebar_expanded
            
            # Actualizar perfil
            profile_info.visible = self.sidebar_expanded
            
            self.page.update()
        
        # Logo y textos del header
        logo_text = ft.Text("IRP--Tools", size=24, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE)
        subtitle_text = ft.Text("Kiting-APP", size=14, color=ft.Colors.WHITE54)
        
        # Información del perfil
        profile_info = ft.Column([
            ft.Text(f"{self.current_user['first_name']} {self.current_user['last_name']}", 
                   size=14, weight=ft.FontWeight.W_500),
            ft.Text(self.current_user['role'].capitalize(), size=12, color=ft.Colors.WHITE54),
        ], spacing=2)
        
        # Crear items de navegación
        nav_container = ft.Column(spacing=5)
        for item in nav_items:
            if item:
                nav_item = create_nav_item(item["icon"], item["label"], item["index"], item["permission"])
                if nav_item:
                    nav_container.controls.append(nav_item)
        
        # Sidebar con hover
        sidebar = ft.Container(
            width=250,
            bgcolor=ft.Colors.GREY_900,
            content=ft.Column([
                # Logo/Header
                ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.Icons.BUSINESS, color=self.theme_color, size=40),
                        logo_text,
                        subtitle_text,
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=ft.padding.symmetric(vertical=30),
                ),
                ft.Divider(height=1, color=ft.Colors.WHITE12),
                # Navigation items
                ft.Container(
                    content=nav_container,
                    padding=ft.padding.only(top=20),
                ),
                ft.Container(expand=True),
                ft.Container(
                    content=ft.Row([
                        ft.IconButton(
                            icon=ft.Icons.LOGOUT,
                            tooltip="Cerrar Sesión",
                            on_click=lambda _: self.logout(),
                        ) if not self.sidebar_expanded else
                        ft.TextButton(
                            "Cerrar Sesión",
                            icon=ft.Icons.LOGOUT,
                            on_click=lambda _: self.logout(),
                        )
                    ]),
                    padding=10,
                ),
                # User profile
                ft.Container(
                    content=ft.Row([
                        ft.CircleAvatar(
                            content=ft.Text(f"{self.current_user['first_name'][0]}{self.current_user['last_name'][0]}"),
                            bgcolor=self.theme_color,
                            radius=20,
                        ),
                        profile_info,
                    ], spacing=15),
                    padding=20,
                ),
            ]),
            on_hover=toggle_sidebar,
            animate=ft.Animation(200, ft.AnimationCurve.EASE_IN_OUT),
        )
        
        # Función para crear contenido según el tab
        def create_content(index):
            if index == 0:  # Dashboard
                return create_dashboard_content()
            elif index == 1:  # Analytics
                return create_analytics_content()
            elif index == 2:  # Team
                return create_team_content()
            elif index == 3:  # Finance
                return create_finance_content()
            elif index == 4:  # Settings
                return create_settings_content()
            elif index == 5:  # Admin
                return create_admin_content()
            elif index == 6:  # Usuarios
                return create_users_management_content()
        
        # Dashboard Content
        def create_dashboard_content():
            return ft.Container(
                content=ft.Column([
                    # Header
                    ft.Container(
                        content=ft.Row([
                            ft.Column([
                                ft.Text(f"Welcome back, {self.current_user['first_name']}", 
                                       size=32, weight=ft.FontWeight.BOLD),
                                ft.Text(f"Today is {datetime.now().strftime('%B %d, %Y')}", 
                                    size=16, color=ft.Colors.WHITE54),
                            ]),
                            ft.Container(expand=True),
                            ft.ElevatedButton(
                                text="Generate Report",
                                icon=ft.Icons.DOWNLOAD,
                                bgcolor=self.theme_color,
                                color=ft.Colors.WHITE,
                                on_click=self.on_generate_report
                            )
                        ]),
                        padding=ft.padding.only(bottom=30),
                    ),
                    # KPI Cards
                    ft.Row([
                        create_kpi_card("Revenue", "$2.4M", "+12%", ft.Icons.TRENDING_UP, ft.Colors.GREEN),
                        create_kpi_card("Users", "48.2K", "+8%", ft.Icons.PEOPLE, ft.Colors.BLUE),
                        create_kpi_card("Conversion", "3.2%", "-2%", ft.Icons.SWAP_VERT, ft.Colors.RED),
                        create_kpi_card("Profit", "$890K", "+24%", ft.Icons.ATTACH_MONEY, ft.Colors.GREEN),
                    ], spacing=20),
                    ft.Container(height=30),
                    # Charts
                    ft.Row([
                        create_chart_container("Monthly Revenue", create_bar_chart()),
                        create_chart_container("User Growth", create_line_chart()),
                    ], spacing=20, expand=True),
                ]),
                padding=40,
            )
    
        # Analytics Content
        def create_analytics_content():
            return ft.Container(
                content=ft.Column([
                    ft.Text("Analytics Overview", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    ft.Row([
                        create_metric_card("Bounce Rate", "42.3%", ft.Colors.ORANGE),
                        create_metric_card("Session Duration", "3m 42s", ft.Colors.PURPLE),
                        create_metric_card("Page Views", "1.2M", ft.Colors.CYAN),
                    ], spacing=20),
                    ft.Container(height=30),
                    create_data_table(),
                ]),
                padding=40,
            )
        
        # Team Content
        def create_team_content():
            team_members = [
                {"name": "Alice Johnson", "role": "CTO", "status": "Active"},
                {"name": "Bob Smith", "role": "CFO", "status": "In Meeting"},
                {"name": "Carol Williams", "role": "CMO", "status": "Active"},
                {"name": "David Brown", "role": "COO", "status": "Away"},
            ]
            
            return ft.Container(
                content=ft.Column([
                    ft.Text("Team Management", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    ft.Row([
                        create_team_card(member) for member in team_members
                    ], wrap=True, spacing=20),
                ]),
                padding=40,
            )
        
        # Finance Content (Solo admin)
        def create_finance_content():
            if self.current_user['role'] != 'admin':
                return ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.Icons.LOCK, size=80, color=ft.Colors.RED_400),
                        ft.Text("Acceso Denegado", size=24, weight=ft.FontWeight.BOLD),
                        ft.Text("No tienes permisos para ver esta sección", color=ft.Colors.WHITE54),
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=40,
                    alignment=ft.alignment.center,
                )
            
            return ft.Container(
                content=ft.Column([
                    ft.Text("Financial Overview", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    create_finance_summary(),
                ]),
                padding=40,
            )
        
        # Settings Content
        def create_settings_content():
            return ft.Container(
                content=ft.Column([
                    ft.Text("Settings", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    create_settings_panel(),
                ]),
                padding=40,
            )
        
        # Admin Content simplificado - solo dashboards externos
        def create_admin_content():
            return ft.Container(
                content=ft.Column([
                    ft.Text("Panel de Administración", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    
                    # Sección de herramientas especiales - DASHBOARDS EXTERNOS
                    ft.Container(
                        content=ft.Column([
                            ft.Text("🛠️ Herramientas Especiales", size=20, weight=ft.FontWeight.W_500),
                            ft.Text("Dashboards externos disponibles:", size=14, color=ft.Colors.WHITE70),
                            ft.Container(height=30),
                            ft.Row([
                                ft.ElevatedButton(
                                    "🎯 Credit Memo Dashboard",
                                    icon=ft.Icons.ANALYTICS,
                                    bgcolor=ft.Colors.GREEN_700,
                                    color=ft.Colors.WHITE,
                                    height=80,
                                    width=250,
                                    on_click=self.launch_credit_memo_dashboard,
                                    style=ft.ButtonStyle(
                                        shape=ft.RoundedRectangleBorder(radius=15)
                                    )
                                ),
                                ft.ElevatedButton(
                                    "📦 Grupos/Kiteo Dashboard",
                                    icon=ft.Icons.INVENTORY,
                                    bgcolor=ft.Colors.PURPLE_700,
                                    color=ft.Colors.WHITE,
                                    height=80,
                                    width=250,
                                    on_click=self.launch_kiteo_dashboard,
                                    style=ft.ButtonStyle(
                                        shape=ft.RoundedRectangleBorder(radius=15)
                                    )
                                ),
                                ft.ElevatedButton(
                                    "🔧 Función de Referencia",
                                    icon=ft.Icons.INFO_OUTLINE,
                                    bgcolor=ft.Colors.GREY_600,
                                    color=ft.Colors.WHITE,
                                    height=80,
                                    width=250,
                                    on_click=self.launch_vew2_dashboard,
                                    style=ft.ButtonStyle(
                                        shape=ft.RoundedRectangleBorder(radius=15)
                                    )
                                ),
                            ], spacing=20, wrap=True),
                            ft.Container(height=20),
                            ft.Text("📁 Archivos en directorio raíz:", size=12, color=ft.Colors.WHITE54, italic=True),
                            ft.Text("• cm.py (Credit Memos) ✅", size=11, color=ft.Colors.GREEN_400),
                            ft.Text("• grupo.py (Grupos/Kiteo) ✅", size=11, color=ft.Colors.PURPLE_400),
                            ft.Text("• Función de referencia (Sin implementar)", size=11, color=ft.Colors.GREY_400),
                            ft.Container(height=20),
                            ft.Text("💡 Los dashboards se ejecutan en procesos independientes", 
                                   size=12, color=ft.Colors.BLUE_400, italic=True),
                        ]),
                        padding=40,
                        bgcolor=ft.Colors.BLUE_GREY_800,
                        border_radius=20,
                    ),
                ]),
                padding=40,
            )
        
        # Nueva sección de gestión de usuarios mejorada
        def create_users_management_content():
            users_list = ft.ListView(spacing=10, expand=True)
            selected_user = [None]  # Para mantener el usuario seleccionado
            
            # Campos de edición
            edit_employee_id = ft.TextField(label="ID Empleado", width=150, disabled=True)
            edit_first_name = ft.TextField(label="Nombre", width=200)
            edit_last_name = ft.TextField(label="Apellido", width=200)
            edit_password = ft.TextField(label="Contraseña", width=200, password=True, can_reveal_password=True)
            edit_role = ft.Dropdown(
                label="Rol",
                width=150,
                options=[
                    ft.dropdown.Option("admin", "Admin"),
                    ft.dropdown.Option("production", "Producción"),
                ],
            )
            
            # Campos para nuevo usuario
            new_employee_id = ft.TextField(label="ID Empleado", width=150, keyboard_type=ft.KeyboardType.NUMBER)
            new_first_name = ft.TextField(label="Nombre", width=200)
            new_last_name = ft.TextField(label="Apellido", width=200)
            new_role = ft.Dropdown(
                label="Rol",
                width=150,
                options=[
                    ft.dropdown.Option("admin", "Administrador"),
                    ft.dropdown.Option("production", "Producción"),
                ],
                value="production",
            )
            
            generated_password = ft.Text("", size=14, color=ft.Colors.GREEN_400, visible=False)
            edit_section = ft.Container(visible=False)
            
            def refresh_users():
                users_list.controls.clear()
                users = self.db.get_all_users()
                
                for user in users:
                    # Obtener contraseña real para mostrar
                    conn = sqlite3.connect(self.db.db_path)
                    cursor = conn.cursor()
                    cursor.execute("SELECT password FROM users WHERE employee_id = ?", (user['employee_id'],))
                    password_result = cursor.fetchone()
                    display_password = password_result[0] if password_result else "N/A"
                    conn.close()
                    
                    # Si es admin default, mostrar contraseña sin hash
                    if user['employee_id'] == 1000:
                        display_password = "AD99"
                    
                    user_card = ft.Container(
                        content=ft.Row([
                            ft.Icon(
                                ft.Icons.ADMIN_PANEL_SETTINGS if user['role'] == 'admin' else ft.Icons.PERSON,
                                size=40, 
                                color=ft.Colors.GOLD if user['role'] == 'admin' else self.theme_color
                            ),
                            ft.Column([
                                ft.Text(f"{user['first_name']} {user['last_name']}", 
                                       size=16, weight=ft.FontWeight.W_500),
                                ft.Text(f"ID: {user['employee_id']} | {user['role'].upper()}", 
                                       size=12, color=ft.Colors.WHITE70),
                                ft.Text(f"Contraseña: {display_password}", 
                                       size=12, color=ft.Colors.CYAN_400),
                                ft.Text(f"Creado: {user['created_at'][:16]}", 
                                       size=10, color=ft.Colors.WHITE54),
                            ], expand=True, spacing=2),
                            ft.Column([
                                ft.ElevatedButton(
                                    "Editar",
                                    icon=ft.Icons.EDIT,
                                    bgcolor=ft.Colors.BLUE_600,
                                    color=ft.Colors.WHITE,
                                    height=35,
                                    on_click=lambda _, u=user: select_user_for_edit(u),
                                ),
                                ft.ElevatedButton(
                                    "Eliminar",
                                    icon=ft.Icons.DELETE,
                                    bgcolor=ft.Colors.RED_600,
                                    color=ft.Colors.WHITE,
                                    height=35,
                                    on_click=lambda _, uid=user['employee_id']: delete_user(uid),
                                    disabled=user['employee_id'] == 1000,
                                ),
                            ], spacing=5),
                        ]),
                        padding=20,
                        bgcolor=ft.Colors.GREY_850 if user['role'] == 'admin' else ft.Colors.GREY_800,
                        border=ft.border.all(2, ft.Colors.GOLD if user['role'] == 'admin' else ft.Colors.TRANSPARENT),
                        border_radius=15,
                        margin=ft.margin.only(bottom=10),
                    )
                    users_list.controls.append(user_card)
                
                self.page.update()
            
            def select_user_for_edit(user):
                selected_user[0] = user
                edit_employee_id.value = str(user['employee_id'])
                edit_first_name.value = user['first_name']
                edit_last_name.value = user['last_name']
                edit_role.value = user['role']
                
                # Obtener contraseña real
                conn = sqlite3.connect(self.db.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT password FROM users WHERE employee_id = ?", (user['employee_id'],))
                password_result = cursor.fetchone()
                edit_password.value = password_result[0] if password_result else ""
                conn.close()
                
                if user['employee_id'] == 1000:
                    edit_password.value = "AD99"
                
                edit_section.visible = True
                self.page.update()
            
            def save_user_changes(e):
                if not selected_user[0]:
                    return
                
                try:
                    conn = sqlite3.connect(self.db.db_path)
                    cursor = conn.cursor()
                    
                    # Actualizar usuario
                    password_to_save = edit_password.value
                    if selected_user[0]['employee_id'] != 1000:
                        # Para usuarios normales, hashear la contraseña
                        password_to_save = self.db.hash_password(edit_password.value)
                    
                    cursor.execute('''
                        UPDATE users 
                        SET first_name = ?, last_name = ?, password = ?, role = ?
                        WHERE employee_id = ?
                    ''', (edit_first_name.value, edit_last_name.value, password_to_save, 
                          edit_role.value, selected_user[0]['employee_id']))
                    
                    conn.commit()
                    conn.close()
                    
                    self.show_snackbar_simple("✅ Usuario actualizado correctamente", ft.Colors.GREEN_700)
                    edit_section.visible = False
                    selected_user[0] = None
                    refresh_users()
                    
                except Exception as ex:
                    self.show_snackbar_simple(f"❌ Error: {str(ex)}", ft.Colors.RED_400)
            
            def cancel_edit(e):
                edit_section.visible = False
                selected_user[0] = None
                self.page.update()
            
            def delete_user(employee_id):
                if self.db.delete_user(employee_id):
                    self.show_snackbar_simple("✅ Usuario eliminado", ft.Colors.GREEN_700)
                    refresh_users()
                else:
                    self.show_snackbar_simple("❌ No se puede eliminar el último administrador", ft.Colors.RED_400)
            
            def create_user(e):
                try:
                    if not all([new_employee_id.value, new_first_name.value, new_last_name.value]):
                        self.show_snackbar_simple("⚠️ Complete todos los campos", ft.Colors.ORANGE_400)
                        return
                    
                    password = self.db.create_user(
                        int(new_employee_id.value),
                        new_first_name.value,
                        new_last_name.value,
                        new_role.value
                    )
                    
                    generated_password.value = f"✅ Usuario creado. Contraseña: {password}"
                    generated_password.visible = True
                    
                    # Limpiar campos
                    new_employee_id.value = ""
                    new_first_name.value = ""
                    new_last_name.value = ""
                    new_role.value = "production"
                    
                    refresh_users()
                    
                except Exception as ex:
                    self.show_snackbar_simple(f"❌ {str(ex)}", ft.Colors.RED_400)
            
            # Crear sección de edición
            edit_section.content = ft.Column([
                ft.Text("✏️ Editar Usuario", size=18, weight=ft.FontWeight.W_500, color=ft.Colors.BLUE_400),
                ft.Container(height=10),
                ft.Row([edit_employee_id, edit_first_name, edit_last_name, edit_role], spacing=15),
                ft.Container(height=10),
                edit_password,
                ft.Container(height=15),
                ft.Row([
                    ft.ElevatedButton(
                        "💾 Guardar Cambios",
                        icon=ft.Icons.SAVE,
                        bgcolor=ft.Colors.GREEN_600,
                        color=ft.Colors.WHITE,
                        on_click=save_user_changes,
                    ),
                    ft.ElevatedButton(
                        "❌ Cancelar",
                        icon=ft.Icons.CANCEL,
                        bgcolor=ft.Colors.RED_600,
                        color=ft.Colors.WHITE,
                        on_click=cancel_edit,
                    ),
                ], spacing=15),
            ])
            
            refresh_users()
            
            return ft.Container(
                content=ft.Column([
                    ft.Text("👥 Gestión de Usuarios", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=20),
                    
                    # Sección de crear usuario
                    ft.Container(
                        content=ft.Column([
                            ft.Text("➕ Crear Nuevo Usuario", size=20, weight=ft.FontWeight.W_500, color=ft.Colors.GREEN_400),
                            ft.Container(height=15),
                            ft.Row([new_employee_id, new_first_name, new_last_name, new_role], spacing=15),
                            ft.Container(height=15),
                            ft.Row([
                                ft.ElevatedButton(
                                    "👤 Crear Usuario",
                                    icon=ft.Icons.PERSON_ADD,
                                    bgcolor=self.theme_color,
                                    color=ft.Colors.WHITE,
                                    height=45,
                                    on_click=create_user,
                                ),
                                generated_password,
                            ], spacing=20),
                        ]),
                        padding=25,
                        bgcolor=ft.Colors.GREEN_900,
                        border_radius=15,
                        border=ft.border.all(1, ft.Colors.GREEN_600),
                    ),
                    
                    ft.Container(height=20),
                    
                    # Sección de edición
                    ft.Container(
                        content=edit_section,
                        padding=25,
                        bgcolor=ft.Colors.BLUE_900,
                        border_radius=15,
                        border=ft.border.all(1, ft.Colors.BLUE_600),
                    ),
                    
                    ft.Container(height=20),
                    
                    # Lista de usuarios
                    ft.Text("📋 Usuarios Registrados", size=20, weight=ft.FontWeight.W_500),
                    ft.Text("💡 Click en 'Editar' para modificar cualquier campo incluida la contraseña", 
                        size=12, color=ft.Colors.WHITE54, italic=True),
                    ft.Container(height=15),
                    ft.Container(
                        content=users_list,
                        padding=10,
                        expand=True,
                        bgcolor=ft.Colors.GREY_900,
                        border_radius=10,
                    ),
                ]),
                padding=40,
            )
        
        # Helper functions para crear componentes
        def create_kpi_card(title, value, change, icon, color):
            return ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(icon, color=color, size=30),
                        ft.Container(expand=True),
                        ft.Container(
                            content=ft.Text(change, color=color, size=14, weight=ft.FontWeight.BOLD),
                            bgcolor=ft.Colors.with_opacity(0.2, color),
                            padding=ft.padding.symmetric(horizontal=10, vertical=5),
                            border_radius=20,
                        ),
                    ]),
                    ft.Container(height=20),
                    ft.Text(value, size=32, weight=ft.FontWeight.BOLD),
                    ft.Text(title, size=16, color=ft.Colors.WHITE54),
                ]),
                bgcolor=ft.Colors.GREY_800,
                padding=25,
                border_radius=15,
                expand=True,
                animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
            )
        
        def create_chart_container(title, chart):
            return ft.Container(
                content=ft.Column([
                    ft.Text(title, size=20, weight=ft.FontWeight.W_500),
                    ft.Container(height=20),
                    chart,
                ]),
                bgcolor=ft.Colors.GREY_800,
                padding=25,
                border_radius=15,
                expand=True,
            )
        
        def create_bar_chart():
            bars = []
            for i in range(12):
                height = random.randint(100, 250)
                bars.append(
                    ft.Container(
                        width=30,
                        height=height,
                        bgcolor=self.theme_color if i == 11 else ft.Colors.BLUE_200,
                        border_radius=ft.border_radius.only(top_left=5, top_right=5),
                    )
                )
            return ft.Row(bars, spacing=10, alignment=ft.MainAxisAlignment.CENTER)
        
        def create_line_chart():
            return ft.Container(
                height=200,
                content=ft.Text("📈 Growth Chart Visualization", size=16, color=ft.Colors.WHITE54),
                alignment=ft.alignment.center,
            )
        
        def create_metric_card(title, value, color):
            return ft.Container(
                content=ft.Column([
                    ft.Text(title, size=14, color=ft.Colors.WHITE54),
                    ft.Text(value, size=28, weight=ft.FontWeight.BOLD, color=color),
                ], spacing=10),
                bgcolor=ft.Colors.GREY_800,
                padding=20,
                border_radius=10,
                expand=True,
            )
        
        def create_data_table():
            return ft.Container(
                content=ft.DataTable(
                    columns=[
                        ft.DataColumn(ft.Text("Metric")),
                        ft.DataColumn(ft.Text("This Month")),
                        ft.DataColumn(ft.Text("Last Month")),
                        ft.DataColumn(ft.Text("Change")),
                    ],
                    rows=[
                        ft.DataRow(cells=[
                            ft.DataCell(ft.Text("Visitors")),
                            ft.DataCell(ft.Text("125,420")),
                            ft.DataCell(ft.Text("115,380")),
                            ft.DataCell(ft.Text("+8.7%", color=ft.Colors.GREEN)),
                        ]),
                        ft.DataRow(cells=[
                            ft.DataCell(ft.Text("Revenue")),
                            ft.DataCell(ft.Text("$458,200")),
                            ft.DataCell(ft.Text("$412,500")),
                            ft.DataCell(ft.Text("+11.1%", color=ft.Colors.GREEN)),
                        ]),
                    ],
                ),
                bgcolor=ft.Colors.GREY_800,
                padding=20,
                border_radius=10,
            )
        
        def create_team_card(member):
            status_color = ft.Colors.GREEN if member["status"] == "Active" else \
                        ft.Colors.ORANGE if member["status"] == "In Meeting" else ft.Colors.GREY
            
            return ft.Container(
                width=250,
                content=ft.Column([
                    ft.Row([
                        ft.CircleAvatar(
                            content=ft.Text(member["name"][0]),
                            bgcolor=self.theme_color,
                            radius=25,
                        ),
                        ft.Column([
                            ft.Text(member["name"], weight=ft.FontWeight.W_500),
                            ft.Text(member["role"], size=14, color=ft.Colors.WHITE54),
                            ft.Row([
                                ft.Container(
                                    width=8,
                                    height=8,
                                    bgcolor=status_color,
                                    border_radius=4,
                                ),
                                ft.Text(member["status"], size=12),
                            ], spacing=5),
                        ], spacing=2),
                    ], spacing=15),
                ]),
                bgcolor=ft.Colors.GREY_800,
                padding=20,
                border_radius=10,
            )
        
        def create_finance_summary():
            return ft.Container(
                content=ft.Column([
                    ft.Row([
                        create_finance_card("Total Assets", "$12.4M", ft.Icons.ACCOUNT_BALANCE),
                        create_finance_card("Liabilities", "$3.2M", ft.Icons.MONEY_OFF),
                        create_finance_card("Net Worth", "$9.2M", ft.Icons.TRENDING_UP),
                    ], spacing=20),
                ]),
            )
        
        def create_finance_card(title, value, icon):
            return ft.Container(
                content=ft.Column([
                    ft.Icon(icon, size=40, color=self.theme_color),
                    ft.Container(height=20),
                    ft.Text(value, size=28, weight=ft.FontWeight.BOLD),
                    ft.Text(title, size=16, color=ft.Colors.WHITE54),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.GREY_800,
                padding=30,
                border_radius=15,
                expand=True,
            )
        
        def create_settings_panel():
            return ft.Container(
                content=ft.Column([
                    create_setting_item("Notifications", True),
                    create_setting_item("Dark Mode", True),
                    create_setting_item("Auto-sync", False),
                    create_setting_item("Analytics", True),
                ], spacing=15),
                bgcolor=ft.Colors.GREY_800,
                padding=30,
                border_radius=15,
                width=500,
            )
        
        def create_setting_item(label, value):
            return ft.Row([
                ft.Text(label, size=16),
                ft.Container(expand=True),
                ft.Switch(value=value, active_color=self.theme_color),
            ])
        
        # Content container
        content_container = ft.Container(
            content=create_content(0),
            expand=True,
            bgcolor=ft.Colors.BLACK,
            animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
        )
        
        # Main layout
        self.page.add(
            ft.Row([
                sidebar,
                content_container,
            ], spacing=0, expand=True)
        )

    def on_generate_report(self, e):
        print(f"Generando reporte.{self.current_user['first_name']} {self.current_user['last_name']}..")
        self.show_snackbar_simple("📄 Reporte generado exitosamente", ft.Colors.GREEN_700)


# Run the app
if __name__ == "__main__":
    app = ExecutiveDashboard()
    ft.app(target=app.main)