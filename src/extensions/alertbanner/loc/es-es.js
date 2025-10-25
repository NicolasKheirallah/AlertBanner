define([], function() {
  return {
    "Title": "AlertBannerApplicationCustomizer",
    
    // General UI
    "Save": "Guardar",
    "Cancel": "Cancelar",
    "Close": "Cerrar",
    "Edit": "Editar",
    "Delete": "Eliminar",
    "Add": "Añadir",
    "Update": "Actualizar",
    "Create": "Crear",
    "Remove": "Quitar",
    "Yes": "Sí",
    "No": "No",
    "OK": "OK",
    "Loading": "Cargando...",
    "Error": "Error",
    "Success": "Éxito",
    "Warning": "Advertencia",
    "Info": "Información",
    
    // Language
    "Language": "Idioma",
    "SelectLanguage": "Seleccionar idioma",
    "ChangeLanguage": "Cambiar idioma",
    
    // Time-related
    "JustNow": "Ahora mismo",
    "MinutesAgo": "Hace {0} minutos",
    "HoursAgo": "Hace {0} horas",
    "DaysAgo": "Hace {0} días",
    
    // Alert Settings
    "AlertSettings": "Configuración de alertas",
    "AlertSettingsTitle": "Configuración del banner de alertas",
    "AlertSettingsDescription": "Configure los ajustes del banner de alertas. Estos cambios se aplicarán en todo el sitio.",
    "ConfigureAlertBannerSettings": "Configurar ajustes del banner de alertas",
    "Features": "Características",
    "EnableUserTargeting": "Habilitar segmentación de usuarios",
    "EnableUserTargetingDescription": "Permitir que las alertas se dirijan a usuarios o grupos específicos",
    "EnableNotifications": "Habilitar notificaciones",
    "EnableNotificationsDescription": "Enviar notificaciones del navegador para alertas críticas",
    "EnableRichMedia": "Habilitar medios enriquecidos",
    "EnableRichMediaDescription": "Soporte para imágenes, videos y contenido enriquecido en alertas",
    "AlertTypesConfiguration": "Configuración de tipos de alerta",
    "AlertTypesConfigurationDescription": "Configure los tipos de alerta disponibles (formato JSON):",
    "AlertTypesPlaceholder": "Ingrese la configuración JSON de tipos de alerta...",
    "AlertTypesHelpText": "Cada tipo de alerta debe tener: nombre, nombreIcono, colorFondo, colorTexto, estilosAdicionales y estilosPrioridad",
    "SaveSettings": "Guardar configuración",
    "InvalidJSONError": "Formato JSON no válido en la configuración de tipos de alerta. Por favor revise su sintaxis.",
    "ManageAlerts": "Manage Alerts",
    "AlertTypesTabTitle": "Alert Types",
    "SettingsTabTitle": "Settings",

    // Create Alert Form
    "CreateAlertSectionContentClassificationTitle": "Content Classification",
    "ContentTypeLabel": "Content Type",
    "CreateAlertSectionContentClassificationDescription": "Choose how this alert content should be classified.",
    "CreateAlertLanguageConfigurationLabel": "Language Configuration",
    "CreateAlertSingleLanguageButton": "Single Language",
    "CreateAlertMultiLanguageButton": "Multi-language",
    "CreateAlertSectionLanguageTargetingTitle": "Language Targeting",
    "CreateAlertTargetLanguageLabel": "Target Language",
    "CreateAlertSectionLanguageTargetingDescription": "Select the language that should receive this alert.",
    "CreateAlertSectionBasicInformationTitle": "Basic Information",
    "CreateAlertTitlePlaceholder": "Enter alert title",
    "CreateAlertTitleDescription": "Provide a short, descriptive title for your alert.",
    "CreateAlertDescriptionPlaceholder": "Enter alert details",
    "CreateAlertDescriptionHelp": "Use rich text to provide the full alert message. You can include links and formatting.",
    "CreateAlertPriorityLowDescription": "Actualizaciones generales o elementos de baja urgencia.",
    "CreateAlertPriorityMediumDescription": "Información importante que debe verse pronto.",
    "CreateAlertPriorityHighDescription": "Alerta sensible que requiere atención rápida.",
    "CreateAlertPriorityCriticalDescription": "Incidente crítico o caída que necesita acción inmediata.",
    "CreateAlertNotificationNoneDescription": "No enviar notificación; los usuarios solo verán la banner.",
    "CreateAlertNotificationBrowserDescription": "Enviar una notificación del navegador junto con la banner.",
    "CreateAlertNotificationEmailDescription": "Enviar un correo electrónico a los usuarios objetivo cuando se publique la alerta.",
    "CreateAlertNotificationBothDescription": "Enviar notificaciones tanto del navegador como por correo electrónico.",
    "CreateAlertContentTypeAlertDescription": "📢 Alerta - Contenido en vivo para usuarios",
    "CreateAlertContentTypeTemplateDescription": "📄 Plantilla - Plantilla reutilizable para alertas futuras",
    "CreateAlertContentTypeDraftDescription": "✏️ Borrador - Trabajo en progreso",
    "CreateAlertTargetLanguageAll": "Todos los idiomas compatibles",
    
    // Alert Management
    "CreateAlert": "Crear alerta",
    "EditAlert": "Editar alerta",
    "DeleteAlert": "Eliminar alerta",
    "AlertTitle": "Título de la alerta",
    "AlertDescription": "Descripción",
    "AlertType": "Tipo de alerta",
    "Priority": "Prioridad",
    "Status": "Estado",
    "TargetSites": "Sitios objetivo",
    "LinkUrl": "URL del enlace",
    "LinkDescription": "Descripción del enlace",
    "ScheduledStart": "Inicio programado",
    "ScheduledEnd": "Fin programado",
    "IsPinned": "Fijado",
    "NotificationType": "Tipo de notificación",
    
    // Priority Levels
    "PriorityLow": "Bajo",
    "PriorityMedium": "Medio",
    "PriorityHigh": "Alto",
    "PriorityCritical": "Crítico",
    
    // Status Types
    "StatusActive": "Activo",
    "StatusExpired": "Expirado",
    "StatusScheduled": "Programado",
    "StatusInactive": "Inactivo",
    
    // Notification Types
    "NotificationNone": "Ninguna",
    "NotificationBrowser": "Navegador",
    "NotificationEmail": "Correo electrónico",
    "NotificationBoth": "Ambos",
    
    // Alert Types
    "AlertTypeInfo": "Información",
    "AlertTypeWarning": "Advertencia",
    "AlertTypeMaintenance": "Mantenimiento",
    "AlertTypeInterruption": "Interrupción",
    
    // User Interface
    "ShowMore": "Mostrar más",
    "ShowLess": "Mostrar menos",
    "ViewDetails": "Ver detalles",
    "Expand": "Expandir",
    "Collapse": "Contraer",
    "Preview": "Vista previa",
    "Templates": "Plantillas",
    "CustomizeColors": "Personalizar colores",
    
    // Site Selection
    "SelectSites": "Seleccionar sitios",
    "CurrentSite": "Sitio actual",
    "AllSites": "Todos los sitios",
    "HubSites": "Sitios hub",
    "RecentSites": "Sitios recientes",
    "FollowedSites": "Sitios seguidos",
    
    // Permissions and Errors
    "InsufficientPermissions": "Permisos insuficientes para realizar esta acción",
    "PermissionDeniedCreateLists": "El usuario no tiene permisos para crear listas de SharePoint",
    "PermissionDeniedAccessLists": "El usuario no tiene permisos para acceder a las listas de SharePoint",
    "ListsNotFound": "Las listas de SharePoint no existen y no se pueden crear",
    "InitializationFailed": "Fallo en la inicialización de la conexión de SharePoint",
    "ConnectionError": "Ocurrió un error de conexión",
    "SaveError": "Ocurrió un error al guardar",
    "LoadError": "Ocurrió un error al cargar los datos",
    
    // User Friendly Messages
    "NoAlertsMessage": "No hay alertas disponibles actualmente",
    "AlertsLoadingMessage": "Cargando alertas...",
    "AlertCreatedSuccess": "Alerta creada exitosamente",
    "AlertUpdatedSuccess": "Alerta actualizada exitosamente",
    "AlertDeletedSuccess": "Alerta eliminada exitosamente",
    "SettingsSavedSuccess": "Configuración guardada exitosamente",
    
    // Date and Time
    "CreatedBy": "Creado por",
    "CreatedOn": "Creado el",
    "LastModified": "Última modificación",
    "Never": "Nunca",
    "Today": "Hoy",
    "Yesterday": "Ayer",
    "Tomorrow": "Mañana",
    
    // Validation Messages
    "FieldRequired": "Este campo es obligatorio",
    "InvalidUrl": "Por favor ingrese una URL válida",
    "InvalidDate": "Por favor ingrese una fecha válida",
    "InvalidEmail": "Por favor ingrese una dirección de correo electrónico válida",
    "TitleTooLong": "El título es muy largo (máximo 255 caracteres)",
    "DescriptionTooLong": "La descripción es muy larga (máximo 2000 caracteres)",
    
    // Rich Media
    "UploadImage": "Subir imagen",
    "RemoveImage": "Quitar imagen",
    "ImageAltText": "Texto alternativo de la imagen",
    "VideoUrl": "URL del video",
    "EmbedCode": "Código de inserción",
    
    // Accessibility
    "CloseDialog": "Cerrar diálogo",
    "OpenSettings": "Abrir configuración",
    "ExpandAlert": "Expandir alerta",
    "CollapseAlert": "Contraer alerta",
    "AlertActions": "Acciones de alerta",
    "PinAlert": "Fijar alerta",
    "UnpinAlert": "Desfijar alerta",

    // Alert UI
    "AlertHeaderPriorityTooltip": "Priority: {0}",
    "AttachmentsHeader": "📎 Attachments ({0})",
    "FileSizeKilobytes": "{0} KB",
    "FileSizeMegabytes": "{0} MB",
    "AlertActionsOpenLink": "Open link",
    "AlertActionsPrevious": "Previous alert",
    "AlertActionsNext": "Next alert",
    "AlertActionsDismiss": "Dismiss alert",
    "AlertActionsHideForever": "Hide alert forever",
    "AlertCarouselCounter": "{0} of {1}",
    "DefaultAlertTypeName": "Default",
    "AlertsLoadErrorTitle": "Unable to load alerts",
    "AlertsLoadErrorFallback": "An error occurred while loading alerts. Please try refreshing the page or contact your administrator if the problem persists."
  }
});
