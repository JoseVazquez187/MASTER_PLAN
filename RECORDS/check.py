import flet as ft
from datetime import datetime
import threading
import time

# ==========================================
# CONFIGURACIÓN GLOBAL Y CONSTANTES
# ==========================================
class AppConfig:
    """Configuración global de la aplicación"""
    APP_TITLE = "Modern Dashboard Template"
    WINDOW_WIDTH = 1400
    WINDOW_HEIGHT = 900
    
    # Colores del tema dark moderno
    PRIMARY_COLOR = "#6366f1"      # Indigo
    SECONDARY_COLOR = "#8b5cf6"    # Purple
    SUCCESS_COLOR = "#10b981"      # Emerald
    WARNING_COLOR = "#f59e0b"      # Amber
    ERROR_COLOR = "#ef4444"        # Red
    
    # Colores de fondo
    BG_PRIMARY = "#0f172a"         # Slate 900
    BG_SECONDARY = "#1e293b"       # Slate 800
    BG_CARD = "#334155"            # Slate 700
    BG_HOVER = "#475569"           # Slate 600
    
    # Colores de texto
    TEXT_PRIMARY = "#f8fafc"       # Slate 50
    TEXT_SECONDARY = "#cbd5e1"     # Slate 300
    TEXT_MUTED = "#94a3b8"         # Slate 400

# ==========================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ==========================================
class ModernApp:
    def __init__(self):
        # Referencias de la página
        self.page = None
        
        # Referencias de widgets principales
        self.sidebar = None
        self.main_content = None
        self.header = None
        self.footer = None
        
        # Estado de la aplicación
        self.current_view = "dashboard"
        self.is_sidebar_expanded = True
        
        # Referencias de widgets específicos
        self.clock_text = None
        self.status_indicator = None
        self.notification_area = None
        
        # Datos de la aplicación
        self.user_data = {"name": "Usuario Admin", "role": "Administrator"}
        self.app_stats = {"users": 0, "files": 0, "processes": 0}

    # ==========================================
    # INICIALIZACIÓN PRINCIPAL
    # ==========================================
    def main(self, page: ft.Page):
        """Punto de entrada principal de la aplicación"""
        self.page = page
        self.setup_page_config()
        self.build_interface()
        self.start_background_tasks()

    def setup_page_config(self):
        """Configuración inicial de la página"""
        self.page.title = AppConfig.APP_TITLE
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.window_width = AppConfig.WINDOW_WIDTH
        self.page.window_height = AppConfig.WINDOW_HEIGHT
        self.page.window_resizable = True
        self.page.bgcolor = AppConfig.BG_PRIMARY
        self.page.window_title_bar_hidden = False
        self.page.window_title_bar_buttons_hidden = False

    # ==========================================
    # CONSTRUCCIÓN DE LA INTERFAZ
    # ==========================================
    def build_interface(self):
        """Construye toda la interfaz de usuario"""
        # Crear componentes principales
        self.header = self.create_header()
        self.sidebar = self.create_sidebar()
        self.main_content = self.create_main_content()
        self.footer = self.create_footer()
        
        # Layout principal con sidebar y contenido
        main_layout = ft.Row([
            self.sidebar,
            ft.VerticalDivider(width=1, color=AppConfig.BG_CARD),
            ft.Column([
                self.header,
                ft.Divider(height=1, color=AppConfig.BG_CARD),
                self.main_content,
                ft.Divider(height=1, color=AppConfig.BG_CARD),
                self.footer,
            ], expand=True)
        ], expand=True, spacing=0)
        
        # Agregar a la página
        self.page.add(main_layout)

    def create_header(self):
        """Crea la barra de header superior"""
        # Reloj en tiempo real
        self.clock_text = ft.Text(
            "00:00:00",
            size=16,
            weight=ft.FontWeight.W_500,
            color=AppConfig.TEXT_SECONDARY
        )
        
        # Indicador de estado
        self.status_indicator = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.CIRCLE, size=12, color=AppConfig.SUCCESS_COLOR),
                ft.Text("Online", size=12, color=AppConfig.TEXT_SECONDARY)
            ], spacing=5),
            padding=ft.padding.symmetric(horizontal=10, vertical=5),
            bgcolor=AppConfig.BG_CARD,
            border_radius=15
        )
        
        # Área de notificaciones
        self.notification_area = ft.Text(
            "",
            size=14,
            color=AppConfig.SUCCESS_COLOR,
            weight=ft.FontWeight.W_500
        )
        
        return ft.Container(
            content=ft.Row([
                # Título y toggle del sidebar
                ft.Row([
                    ft.IconButton(
                        icon=ft.Icons.MENU,
                        icon_color=AppConfig.TEXT_PRIMARY,
                        on_click=self.toggle_sidebar
                    ),
                    ft.Text(
                        "Dashboard Moderno",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=AppConfig.TEXT_PRIMARY
                    )
                ]),
                
                # Área central para notificaciones
                ft.Container(
                    content=self.notification_area,
                    expand=True,
                    alignment=ft.alignment.center
                ),
                
                # Área derecha con controles
                ft.Row([
                    self.status_indicator,
                    self.clock_text,
                    ft.IconButton(
                        icon=ft.Icons.NOTIFICATIONS,
                        icon_color=AppConfig.TEXT_SECONDARY,
                        on_click=self.show_notifications
                    ),
                    ft.IconButton(
                        icon=ft.Icons.SETTINGS,
                        icon_color=AppConfig.TEXT_SECONDARY,
                        on_click=self.show_settings
                    )
                ], spacing=10)
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=20,
            bgcolor=AppConfig.BG_SECONDARY,
            height=70
        )

    def create_sidebar(self):
        """Crea la barra lateral de navegación"""
        # Botones de navegación
        nav_buttons = [
            self.create_nav_button("Dashboard", ft.Icons.DASHBOARD, "dashboard"),
            self.create_nav_button("Archivos", ft.Icons.FOLDER, "files"),
            self.create_nav_button("Procesamiento", ft.Icons.SETTINGS, "processing"),
            self.create_nav_button("Reportes", ft.Icons.ASSESSMENT, "reports"),
            self.create_nav_button("Base de Datos", ft.Icons.STORAGE, "database"),
            self.create_nav_button("Usuarios", ft.Icons.PEOPLE, "users"),
            self.create_nav_button("Configuración", ft.Icons.TUNE, "config"),
        ]
        
        # Perfil de usuario
        user_profile = ft.Container(
            content=ft.Column([
                ft.CircleAvatar(
                    content=ft.Text("U", color=AppConfig.TEXT_PRIMARY, weight=ft.FontWeight.BOLD),
                    bgcolor=AppConfig.PRIMARY_COLOR,
                    radius=25
                ),
                ft.Text(
                    self.user_data["name"],
                    size=14,
                    weight=ft.FontWeight.W_500,
                    color=AppConfig.TEXT_PRIMARY,
                    text_align=ft.TextAlign.CENTER
                ),
                ft.Text(
                    self.user_data["role"],
                    size=12,
                    color=AppConfig.TEXT_MUTED,
                    text_align=ft.TextAlign.CENTER
                )
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8),
            padding=20,
            bgcolor=AppConfig.BG_CARD,
            border_radius=10,
            margin=ft.margin.symmetric(horizontal=10, vertical=10)
        )
        
        return ft.Container(
            content=ft.Column([
                user_profile,
                ft.Divider(color=AppConfig.BG_CARD),
                ft.Column(nav_buttons, spacing=5),
                ft.Container(expand=True),  # Spacer
                ft.Container(
                    content=ft.Text(
                        "v2.0.0",
                        size=10,
                        color=AppConfig.TEXT_MUTED,
                        text_align=ft.TextAlign.CENTER
                    ),
                    padding=10
                )
            ], spacing=0),
            width=250,
            bgcolor=AppConfig.BG_SECONDARY,
            padding=ft.padding.symmetric(vertical=20)
        )

    def create_nav_button(self, label, icon, view_id):
        """Crea un botón de navegación"""
        is_active = self.current_view == view_id
        
        return ft.Container(
            content=ft.Row([
                ft.Icon(
                    icon,
                    size=20,
                    color=AppConfig.TEXT_PRIMARY if is_active else AppConfig.TEXT_SECONDARY
                ),
                ft.Text(
                    label,
                    size=14,
                    weight=ft.FontWeight.W_500 if is_active else ft.FontWeight.NORMAL,
                    color=AppConfig.TEXT_PRIMARY if is_active else AppConfig.TEXT_SECONDARY
                )
            ], spacing=15),
            padding=ft.padding.symmetric(horizontal=20, vertical=12),
            margin=ft.margin.symmetric(horizontal=10, vertical=2),
            bgcolor=AppConfig.BG_CARD if is_active else None,
            border_radius=8,
            on_click=lambda e, vid=view_id: self.navigate_to(vid),
            ink=True
        )

    def create_main_content(self):
        """Crea el área de contenido principal"""
        return ft.Container(
            content=self.get_view_content(),
            expand=True,
            padding=20,
            bgcolor=AppConfig.BG_PRIMARY
        )

    def create_footer(self):
        """Crea el pie de página"""
        return ft.Container(
            content=ft.Row([
                ft.Text(
                    "© 2024 Dashboard Moderno. Todos los derechos reservados.",
                    size=12,
                    color=AppConfig.TEXT_MUTED
                ),
                ft.Container(expand=True),
                ft.Text(
                    "Desarrollado con Flet",
                    size=12,
                    color=AppConfig.TEXT_MUTED
                )
            ]),
            padding=ft.padding.symmetric(horizontal=20, vertical=10),
            bgcolor=AppConfig.BG_SECONDARY,
            height=50
        )

    # ==========================================
    # GESTIÓN DE VISTAS Y NAVEGACIÓN
    # ==========================================
    def get_view_content(self):
        """Retorna el contenido según la vista actual"""
        if self.current_view == "dashboard":
            return self.create_dashboard_view()
        elif self.current_view == "files":
            return self.create_files_view()
        elif self.current_view == "processing":
            return self.create_processing_view()
        elif self.current_view == "reports":
            return self.create_reports_view()
        elif self.current_view == "database":
            return self.create_database_view()
        elif self.current_view == "users":
            return self.create_users_view()
        elif self.current_view == "config":
            return self.create_config_view()
        else:
            return self.create_dashboard_view()

    def create_dashboard_view(self):
        """Vista del dashboard principal"""
        # Tarjetas de estadísticas
        stats_cards = ft.Row([
            self.create_stat_card("Usuarios Activos", "1,234", ft.Icons.PEOPLE, AppConfig.PRIMARY_COLOR),
            self.create_stat_card("Archivos Procesados", "5,678", ft.Icons.FOLDER, AppConfig.SUCCESS_COLOR),
            self.create_stat_card("Procesos en Ejecución", "12", ft.Icons.PLAY_CIRCLE, AppConfig.WARNING_COLOR),
            self.create_stat_card("Errores Detectados", "3", ft.Icons.ERROR, AppConfig.ERROR_COLOR),
        ], spacing=20)
        
        # Botones de acción rápida
        quick_actions = ft.Column([
            ft.Text(
                "Acciones Rápidas",
                size=18,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Row([
                self.create_action_button("Cargar Archivo", ft.Icons.UPLOAD_FILE, self.upload_file),
                self.create_action_button("Ejecutar Proceso", ft.Icons.PLAY_ARROW, self.run_process),
                self.create_action_button("Ver Reportes", ft.Icons.ANALYTICS, self.view_reports),
                self.create_action_button("Configurar", ft.Icons.SETTINGS, self.open_settings),
            ], spacing=15, wrap=True)
        ], spacing=15)
        
        # Tabla de actividad reciente
        recent_activity = self.create_activity_table()
        
        return ft.Column([
            stats_cards,
            ft.Divider(height=30, color="transparent"),
            quick_actions,
            ft.Divider(height=30, color="transparent"),
            recent_activity,
        ], spacing=0, scroll=ft.ScrollMode.AUTO)

    def create_files_view(self):
        """Vista de gestión de archivos"""
        return ft.Column([
            ft.Text(
                "Gestión de Archivos",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Text(
                "Aquí puedes gestionar todos los archivos del sistema",
                size=14,
                color=AppConfig.TEXT_SECONDARY
            ),
            ft.Divider(height=30, color="transparent"),
            
            # Área de carga de archivos
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.CLOUD_UPLOAD, size=60, color=AppConfig.PRIMARY_COLOR),
                    ft.Text("Arrastra archivos aquí o haz clic para seleccionar", color=AppConfig.TEXT_SECONDARY),
                    ft.ElevatedButton(
                        "Seleccionar Archivos",
                        icon=ft.Icons.FOLDER_OPEN,
                        on_click=self.select_files,
                        bgcolor=AppConfig.PRIMARY_COLOR,
                        color=AppConfig.TEXT_PRIMARY
                    )
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=20),
                padding=40,
                bgcolor=AppConfig.BG_CARD,
                border_radius=15,
                border=ft.border.all(2, AppConfig.BG_HOVER),
                alignment=ft.alignment.center,
                height=300
            )
        ], spacing=20)

    def create_processing_view(self):
        """Vista de procesamiento"""
        return ft.Column([
            ft.Text(
                "Centro de Procesamiento",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Row([
                self.create_process_card("Validación FUP", "En espera", ft.Icons.CHECK_CIRCLE),
                self.create_process_card("Análisis de Datos", "Ejecutando", ft.Icons.ANALYTICS),
                self.create_process_card("Generación de Reportes", "Completado", ft.Icons.ASSIGNMENT),
            ], spacing=20, wrap=True)
        ], spacing=20)

    def create_reports_view(self):
        """Vista de reportes"""
        return ft.Column([
            ft.Text(
                "Reportes y Análisis",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Text("Aquí se mostrarán los reportes generados", color=AppConfig.TEXT_SECONDARY)
        ], spacing=20)

    def create_database_view(self):
        """Vista de base de datos"""
        return ft.Column([
            ft.Text(
                "Gestión de Base de Datos",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Text("Administra las conexiones y datos", color=AppConfig.TEXT_SECONDARY)
        ], spacing=20)

    def create_users_view(self):
        """Vista de usuarios"""
        return ft.Column([
            ft.Text(
                "Gestión de Usuarios",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Text("Administra usuarios y permisos", color=AppConfig.TEXT_SECONDARY)
        ], spacing=20)

    def create_config_view(self):
        """Vista de configuración"""
        return ft.Column([
            ft.Text(
                "Configuración del Sistema",
                size=24,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.Text("Ajusta la configuración de la aplicación", color=AppConfig.TEXT_SECONDARY)
        ], spacing=20)

    # ==========================================
    # COMPONENTES REUTILIZABLES
    # ==========================================
    def create_stat_card(self, title, value, icon, color):
        """Crea una tarjeta de estadística"""
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(icon, size=30, color=color),
                    ft.Container(expand=True),
                    ft.Text(value, size=24, weight=ft.FontWeight.BOLD, color=AppConfig.TEXT_PRIMARY)
                ]),
                ft.Text(title, size=14, color=AppConfig.TEXT_SECONDARY)
            ], spacing=10),
            padding=20,
            bgcolor=AppConfig.BG_CARD,
            border_radius=15,
            width=250,
            height=120
        )

    def create_action_button(self, label, icon, action):
        """Crea un botón de acción"""
        return ft.ElevatedButton(
            content=ft.Row([
                ft.Icon(icon, size=20),
                ft.Text(label, size=14, weight=ft.FontWeight.W_500)
            ], spacing=10),
            on_click=action,
            bgcolor=AppConfig.BG_CARD,
            color=AppConfig.TEXT_PRIMARY,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=10)
            )
        )

    def create_process_card(self, title, status, icon):
        """Crea una tarjeta de proceso"""
        status_color = {
            "En espera": AppConfig.WARNING_COLOR,
            "Ejecutando": AppConfig.PRIMARY_COLOR,
            "Completado": AppConfig.SUCCESS_COLOR,
            "Error": AppConfig.ERROR_COLOR
        }.get(status, AppConfig.TEXT_SECONDARY)
        
        return ft.Container(
            content=ft.Column([
                ft.Icon(icon, size=40, color=status_color),
                ft.Text(title, size=16, weight=ft.FontWeight.BOLD, color=AppConfig.TEXT_PRIMARY),
                ft.Text(status, size=14, color=status_color)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=10),
            padding=25,
            bgcolor=AppConfig.BG_CARD,
            border_radius=15,
            width=200,
            height=150
        )

    def create_activity_table(self):
        """Crea la tabla de actividad reciente"""
        return ft.Column([
            ft.Text(
                "Actividad Reciente",
                size=18,
                weight=ft.FontWeight.BOLD,
                color=AppConfig.TEXT_PRIMARY
            ),
            ft.DataTable(
                columns=[
                    ft.DataColumn(ft.Text("Hora", color=AppConfig.TEXT_PRIMARY)),
                    ft.DataColumn(ft.Text("Acción", color=AppConfig.TEXT_PRIMARY)),
                    ft.DataColumn(ft.Text("Usuario", color=AppConfig.TEXT_PRIMARY)),
                    ft.DataColumn(ft.Text("Estado", color=AppConfig.TEXT_PRIMARY)),
                ],
                rows=[
                    ft.DataRow([
                        ft.DataCell(ft.Text("10:30", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("Archivo cargado", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("Admin", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("Completado", color=AppConfig.SUCCESS_COLOR)),
                    ]),
                    ft.DataRow([
                        ft.DataCell(ft.Text("10:25", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("Validación iniciada", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("Usuario1", color=AppConfig.TEXT_SECONDARY)),
                        ft.DataCell(ft.Text("En proceso", color=AppConfig.WARNING_COLOR)),
                    ]),
                ],
                bgcolor=AppConfig.BG_CARD,
                border_radius=10
            )
        ], spacing=15)

    # ==========================================
    # MANEJADORES DE EVENTOS
    # ==========================================
    def navigate_to(self, view_id):
        """Navega a una vista específica"""
        self.current_view = view_id
        self.rebuild_sidebar()
        self.rebuild_main_content()
        self.show_notification(f"Navegando a {view_id}")

    def toggle_sidebar(self, e):
        """Alterna la visibilidad del sidebar"""
        self.is_sidebar_expanded = not self.is_sidebar_expanded
        
        # Reconstruir el sidebar con el nuevo estado
        new_sidebar = self.create_sidebar()
        
        # Actualizar el layout principal
        self.rebuild_interface()
        
        status = "expandido" if self.is_sidebar_expanded else "contraído"
        self.show_notification(f"Sidebar {status}")

    def rebuild_interface(self):
        """Reconstruye la interfaz completa"""
        # Limpiar la página
        self.page.controls.clear()
        
        # Recrear la interfaz
        self.build_interface()
        
        # Actualizar la página
        self.page.update()

    def show_notifications(self, e):
        """Muestra las notificaciones"""
        self.show_notification("Mostrando notificaciones")

    def show_settings(self, e):
        """Muestra la configuración"""
        self.navigate_to("config")

    # ==========================================
    # ACCIONES DE BOTONES
    # ==========================================
    def upload_file(self, e):
        """Maneja la carga de archivos"""
        self.show_notification("Función de carga de archivos")

    def run_process(self, e):
        """Ejecuta un proceso"""
        self.show_notification("Ejecutando proceso...")

    def view_reports(self, e):
        """Ve los reportes"""
        self.navigate_to("reports")

    def open_settings(self, e):
        """Abre la configuración"""
        self.navigate_to("config")

    def select_files(self, e):
        """Selecciona archivos"""
        self.show_notification("Selector de archivos")

    # ==========================================
    # UTILIDADES Y HELPERS
    # ==========================================
    def show_notification(self, message):
        """Muestra una notificación temporal"""
        if self.notification_area:
            self.notification_area.value = message
            self.page.update()
            
            # Limpiar notificación después de 3 segundos
            def clear_notification():
                time.sleep(3)
                if self.notification_area:
                    self.notification_area.value = ""
                    try:
                        self.page.update()
                    except:
                        pass
            
            threading.Thread(target=clear_notification, daemon=True).start()

    def rebuild_sidebar(self):
        """Reconstruye el sidebar con el estado actualizado"""
        new_sidebar = self.create_sidebar()
        # Aquí actualizarías el sidebar en la interfaz
        # Por simplicidad, solo mostramos notificación
        pass

    def rebuild_main_content(self):
        """Reconstruye el contenido principal"""
        if self.main_content:
            self.main_content.content = self.get_view_content()
            self.page.update()

    # ==========================================
    # TAREAS EN SEGUNDO PLANO
    # ==========================================
    def start_background_tasks(self):
        """Inicia las tareas en segundo plano"""
        self.start_clock()

    def start_clock(self):
        """Inicia el reloj en tiempo real"""
        def update_clock():
            while True:
                try:
                    if self.clock_text and self.page:
                        current_time = datetime.now().strftime("%H:%M:%S")
                        self.clock_text.value = current_time
                        self.page.update()
                    time.sleep(1)
                except:
                    break
        
        clock_thread = threading.Thread(target=update_clock, daemon=True)
        clock_thread.start()

# ==========================================
# FUNCIÓN PRINCIPAL
# ==========================================
def main(page: ft.Page):
    """Función principal de la aplicación"""
    app = ModernApp()
    app.main(page)

# ==========================================
# PUNTO DE ENTRADA
# ==========================================
if __name__ == "__main__":
    print("=== DASHBOARD MODERNO ===")
    print("Iniciando aplicación...")
    ft.app(target=main, view=ft.AppView.FLET_APP)