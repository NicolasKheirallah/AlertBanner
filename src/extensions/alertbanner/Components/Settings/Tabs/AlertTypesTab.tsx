import * as React from "react";
import {
  Add24Regular,
  Save24Regular,
  Delete24Regular,
  Dismiss24Regular,
  Edit24Regular,
  Code24Regular,
  EyeOff24Regular,
  Search24Regular,
} from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointTextArea,
  SharePointSection,
} from "../../UI/SharePointControls";
import SharePointDialog from "../../UI/SharePointDialog";
import ColorPicker from "../../UI/ColorPicker";
import AlertPreview from "../../UI/AlertPreview";
import { AlertPriority, IAlertType } from "../../Alerts/IAlerts";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { NotificationService } from "../../Services/NotificationService";
import { useFluentDialogs } from "../../Hooks/useFluentDialogs";
import styles from "../AlertSettings.module.scss";
import { logger } from "../../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import {
  ALL_FLUENT_ICON_NAMES,
  getAlertTypeIcon,
  getAlertTypeIconLabel,
} from "../../AlertItem/utils";
import { meetsWCAGAA } from "../../Utils/ColorUtils";

const HEX_COLOR_PATTERN = /^#([a-fA-F0-9]{6}|[a-fA-F0-9]{3})$/;

const createDefaultAlertType = (): IAlertType => ({
  name: "",
  iconName: "Info",
  backgroundColor: "#0078d4",
  textColor: "#ffffff",
  additionalStyles: "",
  priorityStyles: {
    [AlertPriority.Critical]: "border: 2px solid #E81123;",
    [AlertPriority.High]: "border: 1px solid #EA4300;",
    [AlertPriority.Medium]: "",
    [AlertPriority.Low]: "",
  },
});

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return strings.CreateAlertUnknownError;
};

const normalizeName = (value: string): string => value.trim();

export interface IAlertTypesTabProps {
  alertTypes: IAlertType[];
  setAlertTypes: React.Dispatch<React.SetStateAction<IAlertType[]>>;
  newAlertType: IAlertType;
  setNewAlertType: React.Dispatch<React.SetStateAction<IAlertType>>;
  isCreatingType: boolean;
  setIsCreatingType: React.Dispatch<React.SetStateAction<boolean>>;
  alertService: SharePointAlertService;
  context?: any;
}

const AlertTypesTab: React.FC<IAlertTypesTabProps> = ({
  alertTypes,
  setAlertTypes,
  newAlertType,
  setNewAlertType,
  isCreatingType,
  setIsCreatingType,
  alertService,
  context,
}) => {
  const [draggedItem, setDraggedItem] = React.useState<number | null>(null);
  const [editingType, setEditingType] = React.useState<IAlertType | null>(null);
  const [isEditMode, setIsEditMode] = React.useState(false);
  const [showAdvancedStyles, setShowAdvancedStyles] = React.useState(false);
  const [isIconDialogOpen, setIsIconDialogOpen] = React.useState(false);
  const [iconSearchTerm, setIconSearchTerm] = React.useState("");
  const { confirm, dialogs } = useFluentDialogs();
  const notificationService = React.useMemo(
    () => (context ? NotificationService.getInstance(context) : null),
    [context],
  );

  const resetTypeForm = React.useCallback(() => {
    setNewAlertType(createDefaultAlertType());
    setEditingType(null);
    setIsEditMode(false);
    setShowAdvancedStyles(false);
    setIsCreatingType(false);
  }, [setIsCreatingType, setNewAlertType]);

  const validationError = React.useMemo(() => {
    const name = normalizeName(newAlertType.name);
    const iconName = (newAlertType.iconName || "").trim();

    if (!name) {
      return "Type name is required.";
    }

    if (name.length < 2) {
      return "Type name must be at least 2 characters.";
    }

    const duplicate = alertTypes.some((type) => {
      const sameName = normalizeName(type.name).toLowerCase() === name.toLowerCase();
      if (!sameName) {
        return false;
      }

      if (!isEditMode || !editingType) {
        return true;
      }

      return normalizeName(type.name).toLowerCase() !== normalizeName(editingType.name).toLowerCase();
    });

    if (duplicate) {
      return "An alert type with this name already exists.";
    }

    if (!iconName) {
      return "Please select an icon.";
    }

    if (!HEX_COLOR_PATTERN.test(newAlertType.backgroundColor)) {
      return "Background color must be a valid hex value.";
    }

    if (!HEX_COLOR_PATTERN.test(newAlertType.textColor)) {
      return "Text color must be a valid hex value.";
    }

    if (!meetsWCAGAA(newAlertType.backgroundColor, newAlertType.textColor)) {
      return "Background and text colors do not meet WCAG AA contrast.";
    }

    return "";
  }, [alertTypes, editingType, isEditMode, newAlertType]);

  const canSubmit = validationError.length === 0;

  const buildNormalizedType = React.useCallback((): IAlertType => {
    const normalizedName = normalizeName(newAlertType.name);
    return {
      ...newAlertType,
      name: normalizedName,
      iconName: (newAlertType.iconName || "Info").trim() || "Info",
      additionalStyles: (newAlertType.additionalStyles || "").trim(),
    };
  }, [newAlertType]);

  const handleCreateAlertType = React.useCallback(async () => {
    if (!canSubmit) {
      notificationService?.showWarning(validationError, strings.Warning);
      return;
    }

    try {
      const normalizedType = buildNormalizedType();
      const updatedTypes = [...alertTypes, normalizedType];

      await alertService.saveAlertTypes(updatedTypes);
      setAlertTypes(updatedTypes);
      notificationService?.showSuccess("Alert type created.", strings.Success);
      resetTypeForm();
    } catch (error) {
      logger.error("AlertTypesTab", "Error creating alert type", error);
      notificationService?.showError(
        `Failed to create alert type: ${getErrorMessage(error)}`,
        strings.Error,
      );
    }
  }, [
    alertService,
    alertTypes,
    buildNormalizedType,
    canSubmit,
    notificationService,
    resetTypeForm,
    setAlertTypes,
    validationError,
  ]);

  const handleDeleteAlertType = React.useCallback(
    async (index: number) => {
      const typeToDelete = alertTypes[index];

      const shouldDelete = await confirm({
        title: "Delete Alert Type",
        message: `Are you sure you want to delete the alert type \"${typeToDelete.name}\"?`,
        confirmText: strings.Delete,
      });

      if (!shouldDelete) {
        return;
      }

      try {
        const updatedTypes = alertTypes.filter((_, i) => i !== index);
        await alertService.saveAlertTypes(updatedTypes);
        setAlertTypes(updatedTypes);
        notificationService?.showSuccess("Alert type deleted.", strings.Success);
      } catch (error) {
        logger.error("AlertTypesTab", "Error deleting alert type", error);
        notificationService?.showError(
          `Failed to delete alert type: ${getErrorMessage(error)}`,
          strings.Error,
        );
      }
    },
    [alertService, alertTypes, confirm, notificationService, setAlertTypes],
  );

  const handleDragStart = React.useCallback(
    (e: React.DragEvent, index: number) => {
      setDraggedItem(index);
      e.dataTransfer.effectAllowed = "move";
    },
    [],
  );

  const handleDragEnd = React.useCallback(() => {
    setDraggedItem(null);
  }, []);

  const handleDragOver = React.useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
  }, []);

  const handleDrop = React.useCallback(
    async (e: React.DragEvent, dropIndex: number) => {
      e.preventDefault();

      if (draggedItem === null || draggedItem === dropIndex) {
        return;
      }

      const updatedTypes = [...alertTypes];
      const draggedType = updatedTypes[draggedItem];

      updatedTypes.splice(draggedItem, 1);
      const insertIndex = draggedItem < dropIndex ? dropIndex - 1 : dropIndex;
      updatedTypes.splice(insertIndex, 0, draggedType);

      try {
        await alertService.saveAlertTypes(updatedTypes);
        setAlertTypes(updatedTypes);
      } catch (error) {
        logger.error("AlertTypesTab", "Error reordering alert types", error);
        notificationService?.showError(
          `Failed to save reordered alert types: ${getErrorMessage(error)}`,
          strings.Error,
        );
      }

      setDraggedItem(null);
    },
    [alertService, alertTypes, draggedItem, notificationService, setAlertTypes],
  );

  const handleEditAlertType = React.useCallback(
    (alertType: IAlertType) => {
      setEditingType({ ...alertType });
      setNewAlertType({ ...alertType });
      setIsEditMode(true);
      setShowAdvancedStyles(!!alertType.additionalStyles);
      setIsCreatingType(true);
    },
    [setIsCreatingType, setNewAlertType],
  );

  const handleUpdateAlertType = React.useCallback(async () => {
    if (!editingType) {
      return;
    }

    if (!canSubmit) {
      notificationService?.showWarning(validationError, strings.Warning);
      return;
    }

    try {
      const normalizedType = buildNormalizedType();
      const editingNameNormalized = normalizeName(editingType.name).toLowerCase();
      const updatedTypes = alertTypes.map((type) =>
        normalizeName(type.name).toLowerCase() === editingNameNormalized
          ? normalizedType
          : type,
      );

      await alertService.saveAlertTypes(updatedTypes);
      setAlertTypes(updatedTypes);
      notificationService?.showSuccess(strings.AlertUpdatedSuccess, strings.Success);
      resetTypeForm();
    } catch (error) {
      logger.error("AlertTypesTab", "Failed to update alert type", error);
      notificationService?.showError(
        `Failed to update alert type: ${getErrorMessage(error)}`,
        strings.Error,
      );
    }
  }, [
    alertService,
    alertTypes,
    buildNormalizedType,
    canSubmit,
    editingType,
    notificationService,
    resetTypeForm,
    setAlertTypes,
    validationError,
  ]);

  const dialogTitle = isEditMode
    ? `Edit Alert Type - ${editingType?.name || ""}`
    : "Create New Alert Type";

  const filteredIconNames = React.useMemo(() => {
    const query = iconSearchTerm.trim().toLowerCase();
    if (!query) {
      return ALL_FLUENT_ICON_NAMES;
    }

    return ALL_FLUENT_ICON_NAMES.filter((iconName) =>
      iconName.toLowerCase().includes(query),
    );
  }, [iconSearchTerm]);

  return (
    <div className={styles.tabPane}>
      <div className={styles.tabHeader}>
        <div>
          <h3>{strings.AlertTypesTabTitle}</h3>
          <p>Create and customize the visual appearance of alert categories.</p>
        </div>
        <SharePointButton
          variant="primary"
          icon={<Add24Regular />}
          onClick={() => setIsCreatingType(true)}
        >
          Create New Type
        </SharePointButton>
      </div>

      <SharePointSection title="Existing Alert Types">
        <div className={styles.dragDropInstructions}>
          <p>
            <strong>Tip:</strong> Drag and drop alert types to reorder them.
          </p>
        </div>

        <div className={styles.existingTypes}>
          {alertTypes.map((type, index) => (
            <div
              key={type.name}
              className={`${styles.alertTypeCard} ${draggedItem === index ? styles.alertCard : ""}`}
              draggable
              onDragStart={(e) => handleDragStart(e, index)}
              onDragEnd={handleDragEnd}
              onDragOver={handleDragOver}
              onDrop={(e) => void handleDrop(e, index)}
            >
              <div className={styles.dragHandle}>
                <span className={styles.dragIcon}>⋮⋮</span>
                <span className={styles.orderNumber}>#{index + 1}</span>
              </div>

              <div className={styles.alertCardContent}>
                <div className={styles.typeTitleRow}>
                  <span className={styles.typeCardIcon}>
                    {getAlertTypeIcon(type.iconName, AlertPriority.Medium)}
                  </span>
                  <h4>{type.name}</h4>
                </div>
              </div>

              <div className={styles.typePreview}>
                <AlertPreview
                  title={`Sample ${type.name} Alert`}
                  description="This is a preview of how this alert type appears."
                  alertType={type}
                  priority={AlertPriority.Medium}
                  isPinned={false}
                />
              </div>

              <div className={styles.typeActions}>
                <SharePointButton
                  variant="secondary"
                  icon={<Edit24Regular />}
                  onClick={() => handleEditAlertType(type)}
                >
                  {strings.Edit}
                </SharePointButton>
                <SharePointButton
                  variant="danger"
                  icon={<Delete24Regular />}
                  onClick={() => void handleDeleteAlertType(index)}
                >
                  {strings.Delete}
                </SharePointButton>
              </div>
            </div>
          ))}

          {alertTypes.length === 0 && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>🎨</div>
              <h4>No Alert Types</h4>
              <p>Create your first alert type to start customizing alert styling.</p>
              <SharePointButton
                variant="primary"
                icon={<Add24Regular />}
                onClick={() => setIsCreatingType(true)}
              >
                Create First Type
              </SharePointButton>
            </div>
          )}
        </div>
      </SharePointSection>

      <SharePointDialog
        isOpen={isCreatingType}
        onClose={resetTypeForm}
        title={dialogTitle}
        width={760}
        footer={
          <div className={styles.formActions}>
            <SharePointButton
              variant="secondary"
              icon={<Dismiss24Regular />}
              onClick={resetTypeForm}
            >
              {strings.Cancel}
            </SharePointButton>
            <SharePointButton
              variant="primary"
              icon={<Save24Regular />}
              onClick={isEditMode ? () => void handleUpdateAlertType() : () => void handleCreateAlertType()}
              disabled={!canSubmit}
            >
              {isEditMode ? "Update Type" : "Create Type"}
            </SharePointButton>
          </div>
        }
      >
        <div className={styles.typeFormWithPreview}>
          <div className={styles.typeFormColumn}>
            <SharePointInput
              label="Type Name"
              value={newAlertType.name}
              onChange={(value) =>
                setNewAlertType((prev) => ({ ...prev, name: value }))
              }
              placeholder="e.g., Maintenance, Emergency, Update"
              required
              description="A unique name for this alert type"
            />

            <div className={styles.iconSelectorSection}>
              <label className={styles.fieldLabel}>
                {strings.AlertTypeIconLabel}
              </label>
              <div className={styles.iconSelectorRow}>
                <div className={styles.iconSelectionPreview}>
                  <span className={styles.iconSelectorGlyph}>
                    {getAlertTypeIcon(
                      newAlertType.iconName,
                      AlertPriority.Medium,
                    )}
                  </span>
                  <span className={styles.iconSelectionText}>
                    {getAlertTypeIconLabel(newAlertType.iconName || "Info")}
                  </span>
                </div>
                <SharePointButton
                  variant="secondary"
                  icon={<Search24Regular />}
                  onClick={() => {
                    setIconSearchTerm("");
                    setIsIconDialogOpen(true);
                  }}
                >
                  {strings.AlertTypeChooseIconButton}
                </SharePointButton>
              </div>
            </div>

            <SharePointInput
              label={strings.AlertTypeSelectedIconNameLabel}
              value={newAlertType.iconName}
              onChange={(value) =>
                setNewAlertType((prev) => ({ ...prev, iconName: value }))
              }
              placeholder={strings.AlertTypeSelectedIconNamePlaceholder}
              description={strings.AlertTypeSelectedIconNameDescription}
            />

            <div className={styles.colorRow}>
              <ColorPicker
                label="Background Color"
                value={newAlertType.backgroundColor}
                onChange={(color) =>
                  setNewAlertType((prev) => ({
                    ...prev,
                    backgroundColor: color,
                  }))
                }
                description="Main background color for this type"
              />
              <ColorPicker
                label="Text Color"
                value={newAlertType.textColor}
                onChange={(color) =>
                  setNewAlertType((prev) => ({ ...prev, textColor: color }))
                }
                description="Text color (must pass contrast checks)"
              />
            </div>

            <SharePointButton
              variant="secondary"
              icon={showAdvancedStyles ? <EyeOff24Regular /> : <Code24Regular />}
              onClick={() => setShowAdvancedStyles((prev) => !prev)}
            >
              {showAdvancedStyles ? "Hide Advanced Styles" : "Show Advanced Styles"}
            </SharePointButton>

            {showAdvancedStyles && (
              <SharePointTextArea
                label="Custom CSS Styles"
                value={newAlertType.additionalStyles || ""}
                onChange={(value) =>
                  setNewAlertType((prev) => ({
                    ...prev,
                    additionalStyles: value,
                  }))
                }
                placeholder="Additional CSS styles (advanced)"
                rows={3}
                description="Optional custom CSS for advanced styling"
              />
            )}

            {!!validationError && (
              <div className={styles.errorMessage} role="alert">
                {validationError}
              </div>
            )}
          </div>

          <div className={styles.typePreviewColumn}>
            <h4>{strings.Preview}</h4>
            <AlertPreview
              title="Sample Alert Title"
              description="This is how this alert type will appear to users."
              alertType={buildNormalizedType()}
              priority={AlertPriority.Medium}
              isPinned={false}
            />
          </div>
        </div>
      </SharePointDialog>

      <SharePointDialog
        isOpen={isIconDialogOpen}
        onClose={() => setIsIconDialogOpen(false)}
        title={strings.AlertTypeIconDialogTitle}
        width={980}
        footer={
          <div className={styles.formActions}>
            <SharePointButton
              variant="secondary"
              onClick={() => setIsIconDialogOpen(false)}
            >
              {strings.Close}
            </SharePointButton>
          </div>
        }
      >
        <div className={styles.iconDialogSearchRow}>
          <SharePointInput
            label={strings.AlertTypeIconSearchLabel}
            value={iconSearchTerm}
            onChange={setIconSearchTerm}
            placeholder={strings.AlertTypeIconSearchPlaceholder}
          />
          <div className={styles.iconDialogMeta}>
            {strings.AlertTypeIconResultsCount.replace(
              "{0}",
              filteredIconNames.length.toString(),
            )}
          </div>
        </div>

        {filteredIconNames.length === 0 ? (
          <div className={styles.emptyState}>
            <h4>{strings.AlertTypeIconNoResultsTitle}</h4>
            <p>{strings.AlertTypeIconNoResultsDescription}</p>
          </div>
        ) : (
          <div className={styles.iconDialogGrid}>
            {filteredIconNames.map((iconName) => {
              const isSelected =
                (newAlertType.iconName || "").toLowerCase() ===
                iconName.toLowerCase();

              return (
                <button
                  key={iconName}
                  type="button"
                  className={`${styles.iconSelectorButton} ${isSelected ? styles.iconSelectorButtonActive : ""}`}
                  onClick={() => {
                    setNewAlertType((prev) => ({ ...prev, iconName }));
                    setIsIconDialogOpen(false);
                  }}
                  aria-pressed={isSelected}
                  title={iconName}
                >
                  <span className={styles.iconSelectorGlyph}>
                    {getAlertTypeIcon(iconName, AlertPriority.Medium)}
                  </span>
                  <span className={styles.iconSelectorLabel}>
                    {getAlertTypeIconLabel(iconName)}
                  </span>
                </button>
              );
            })}
          </div>
        )}
      </SharePointDialog>
      {dialogs}
    </div>
  );
};

export default AlertTypesTab;
