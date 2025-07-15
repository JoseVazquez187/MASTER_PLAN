import flet as ft
from datetime import datetime
import random
import sqlite3
import os
import hashlib
import time
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
            admin_password = self.hash_password("AD99")  # Admin + 99
            cursor.execute('''
                INSERT INTO users (employee_id, first_name, last_name, password, role)
                VALUES (1000, 'Admin', 'Default', ?, 'admin')
            ''', (admin_password,))
        
        conn.commit()
        conn.close()
    
    def hash_password(self, password: str) -> str:
        """Hash de la contraseña"""
        # return password
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
        
        # hashed_password = self.hash_password(password)
        cursor.execute('''
            SELECT employee_id, first_name, last_name, role 
            FROM users 
            WHERE employee_id = ? AND password = ?
        ''', (employee_id, password))
        
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
            # hashed_password = self.hash_password(password)
            
            cursor.execute('''
                INSERT INTO users (employee_id, first_name, last_name, password, role)
                VALUES (?, ?, ?, ?, ?)
            ''', (employee_id, first_name, last_name, password, role))
            
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
    
    def show_login_screen(self):
        """Mostrar pantalla de login - FUNCIONANDO"""
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
        
        # Agregar opción de admin si el usuario es admin
        if self.current_user['role'] == 'admin':
            nav_items.append({
                "icon": ft.Icons.ADMIN_PANEL_SETTINGS, 
                "label": "Admin", 
                "index": 5, 
                "permission": "admin"
            })
        
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
        
        # Admin Content
        def create_admin_content():
            ft.Container(expand=True),    
            ft.ElevatedButton(
                text="TEST CLICK",
                icon=ft.Icons.DOWNLOAD,
                bgcolor=self.theme_color,
                color=ft.Colors.WHITE,
                on_click=self.on_generate_report
            ),
            users_list = ft.ListView(spacing=10, expand=True)

            def refresh_users():
                users_list.controls.clear()
                users = self.db.get_all_users()
                
                for user in users:
                    user_card = ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.Icons.PERSON, size=40, color=self.theme_color),
                            ft.Column([
                                ft.Text(f"{user['first_name']} {user['last_name']}", 
                                       size=16, weight=ft.FontWeight.W_500),
                                ft.Text(f"ID: {user['employee_id']} - {user['role'].capitalize()}", 
                                       size=14, color=ft.Colors.WHITE54),
                            ], expand=True),
                            ft.IconButton(
                                icon=ft.Icons.DELETE,
                                icon_color=ft.Colors.RED_400,
                                on_click=lambda _, uid=user['employee_id']: delete_user(uid),
                                disabled=user['employee_id'] == 1000,  # No eliminar admin default
                            ),
                        ]),
                        padding=15,
                        bgcolor=ft.Colors.GREY_800,
                        border_radius=10,
                        margin=ft.margin.only(bottom=10),
                    )
                    users_list.controls.append(user_card)
                
                self.page.update()
            
            def delete_user(employee_id):
                if self.db.delete_user(employee_id):
                    refresh_users()
                else:
                    # Mostrar error si no se puede eliminar
                    self.page.show_snack_bar(
                        ft.SnackBar(content=ft.Text("No se puede eliminar el último administrador"))
                    )
            
            # Formulario para crear usuario
            new_employee_id = ft.TextField(label="Número de Empleado", width=200, keyboard_type=ft.KeyboardType.NUMBER)
            new_first_name = ft.TextField(label="Nombre", width=200)
            new_last_name = ft.TextField(label="Apellido", width=200)
            new_role = ft.Dropdown(
                label="Rol",
                width=200,
                options=[
                    ft.dropdown.Option("admin", "Administrador"),
                    ft.dropdown.Option("production", "Producción"),
                ],
                value="production",
            )
            
            generated_password = ft.Text("", size=16, color=ft.Colors.GREEN_400, visible=False)
            
            def create_user(e):
                try:
                    password = self.db.create_user(
                        int(new_employee_id.value),
                        new_first_name.value,
                        new_last_name.value,
                        new_role.value
                    )
                    
                    # Mostrar contraseña generada
                    generated_password.value = f"Contraseña generada: {password}"
                    generated_password.visible = True
                    
                    # Limpiar campos
                    new_employee_id.value = ""
                    new_first_name.value = ""
                    new_last_name.value = ""
                    new_role.value = "production"
                    
                    refresh_users()
                    self.page.update()
                    
                except Exception as ex:
                    self.page.show_snack_bar(
                        ft.SnackBar(content=ft.Text(str(ex), color=ft.Colors.RED_400))
                    )
            
            refresh_users()
            
            return ft.Container(
                content=ft.Column([
                    ft.Text("Administración de Usuarios", size=32, weight=ft.FontWeight.BOLD),
                    ft.Container(height=30),
                    
                    # Formulario de nuevo usuario
                    ft.Container(
                        content=ft.Column([
                            ft.Text("Crear Nuevo Usuario", size=20, weight=ft.FontWeight.W_500),
                            ft.Container(height=20),
                            ft.Row([
                                new_employee_id,
                                new_first_name,
                                new_last_name,
                                new_role,
                            ], spacing=20),
                            ft.Container(height=20),
                            ft.Row([
                                ft.ElevatedButton(
                                    "Crear Usuario",
                                    icon=ft.Icons.PERSON_ADD,
                                    bgcolor=self.theme_color,
                                    color=ft.Colors.WHITE,
                                    on_click=create_user,
                                ),
                                generated_password,
                            ], spacing=20),
                        ]),
                        padding=30,
                        bgcolor=ft.Colors.GREY_800,
                        border_radius=15,
                    ),
                    
                    ft.Container(height=30),
                    
                    # Lista de usuarios
                    ft.Text("Usuarios Registrados", size=20, weight=ft.FontWeight.W_500),
                    ft.Container(height=20),
                    ft.Container(
                        content=users_list,
                        padding=10,
                        expand=True,
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

    def on_generate_report(self,e):
        print(f"Generando reporte.{self.current_user['first_name']} {self.current_user['last_name']}..")
        
        
# Run the app
if __name__ == "__main__":
    app = ExecutiveDashboard()
    ft.app(target=app.main)