import { IUser, ITargetingRule, IPersonField } from "../Alerts/IAlerts";
import { IAlertItem } from "../Alerts/IAlerts";
import {
  MSGraphClientV3,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import StorageService from "./StorageService";
import { logger } from "./LoggerService";

interface IGraphGroup {
  id: string;
  displayName: string;
  [key: string]: any;
}

interface ISharePointGroup {
  Id: number;
  Title: string;
  [key: string]: any;
}

export class UserTargetingService {
  private static instance: UserTargetingService;
  private context: ApplicationCustomizerContext;
  private graphClient: MSGraphClientV3;
  private spGroupIds: number[] = [];
  private currentUser: IUser | null = null;
  private userGroups: string[] = [];
  private userGroupIds: string[] = [];
  private isInitialized: boolean = false;
  private storageService: StorageService;

  private constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
  ) {
    this.graphClient = graphClient;
    this.context = context;
    this.storageService = StorageService.getInstance();
  }

  public static getInstance(
    graphClient: MSGraphClientV3,
    context?: ApplicationCustomizerContext,
  ): UserTargetingService {
    if (!UserTargetingService.instance && context) {
      UserTargetingService.instance = new UserTargetingService(
        graphClient,
        context,
      );
    }
    return UserTargetingService.instance;
  }

  public async initialize(): Promise<void> {
    if (this.isInitialized) return;

    try {
      const userResponse = await this.graphClient
        .api("/me")
        .select("id,displayName,mail,jobTitle,department,userPrincipalName")
        .get();

      this.currentUser = {
        id: userResponse.id,
        displayName: userResponse.displayName,
        email: userResponse.mail,
        jobTitle: userResponse.jobTitle,
        department: userResponse.department,
        userGroups: [],
      };

      this.storageService.setUserId(this.currentUser.id);

      const groups: IGraphGroup[] = [];
      let requestPath = "/me/memberOf?$select=id,displayName&$top=100";

      while (requestPath) {
        const response = await this.graphClient.api(requestPath).get();
        if (Array.isArray(response?.value)) {
          groups.push(...response.value);
        }

        if (response["@odata.nextLink"]) {
          requestPath = response["@odata.nextLink"].replace(
            "https://graph.microsoft.com/v1.0",
            "",
          );
        } else {
          requestPath = "";
        }
      }

      if (groups.length > 0) {
        this.userGroups = groups
          .map((group) => group.displayName)
          .filter(Boolean);
        this.userGroupIds = groups.map((group) => group.id).filter(Boolean);
        this.currentUser.userGroups = this.userGroups;
      }

      try {
        const spGroupsResponse: SPHttpClientResponse =
          await this.context.spHttpClient.get(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups?$select=Id,Title`,
            SPHttpClient.configurations.v1,
          );

        if (spGroupsResponse.ok) {
          const spGroups = await spGroupsResponse.json();
          if (spGroups && spGroups.value) {
            const spGroupsList: ISharePointGroup[] = spGroups.value;
            this.spGroupIds = spGroupsList.map((g) => g.Id);
            const spGroupNames = spGroupsList.map((g) => g.Title);
            this.userGroups.push(...spGroupNames);
            this.currentUser.userGroups = this.userGroups;
          }
        }
      } catch (spError) {
        logger.warn(
          "UserTargetingService",
          "Error fetching SharePoint groups",
          spError,
        );
      }

      this.isInitialized = true;
    } catch (error) {
      logger.error(
        "UserTargetingService",
        "Error initializing user targeting service",
        error,
      );
      this.isInitialized = true;
    }
  }

  public async filterAlertsForCurrentUser(
    alerts: IAlertItem[],
  ): Promise<IAlertItem[]> {
    if (!this.isInitialized) {
      await this.initialize();
    }

    // If no user information is available or initialization failed, show all alerts
    if (!this.currentUser) {
      return alerts;
    }

    return alerts.filter((alert) => {
      if (!alert.targetUsers || alert.targetUsers.length === 0) {
        return true;
      }

      return this.isUserTargeted(alert.targetUsers);
    });
  }

  private isUserTargeted(targetUsers: IPersonField[]): boolean {
    if (!this.currentUser || !targetUsers || targetUsers.length === 0) {
      return false;
    }

    return targetUsers.some((person) => {
      if (person.isGroup) {
        // If it's a group, check if user is member of that group
        return this.isUserInGroup(person);
      } else {
        return this.isCurrentUser(person);
      }
    });
  }

  private evaluateTargetingRule(rule: ITargetingRule): boolean {
    if (!this.currentUser) return false;

    if (rule.targetUsers || rule.targetGroups) {
      return this.evaluatePeopleFieldTargeting(rule);
    }
    else if (rule.audiences) {
      return this.evaluateLegacyTargeting(rule);
    }

    return false;
  }

  private evaluatePeopleFieldTargeting(rule: ITargetingRule): boolean {
    if (!this.currentUser) return false;

    const userMatch =
      rule.targetUsers?.some((person) => this.isCurrentUser(person)) || false;

    // Group targeting: Check if current user belongs to any of the target groups
    const groupMatch =
      rule.targetGroups?.some((group) => this.isUserInGroup(group)) || false;

    switch (rule.operation) {
      case "anyOf":
        return userMatch || groupMatch;

      case "allOf":
        // For allOf with both user and group targeting, require both to match
        if (rule.targetUsers && rule.targetGroups) {
          return userMatch && groupMatch;
        }
        return rule.targetUsers ? userMatch : groupMatch;

      case "noneOf":
        // Show if user doesn't match and is not in any target group
        return !userMatch && !groupMatch;

      default:
        return false;
    }
  }

  private evaluateLegacyTargeting(rule: ITargetingRule): boolean {
    if (!this.currentUser || !rule.audiences) return false;

    const userProperties = [
      ...(this.userGroups || []),
      this.currentUser.department,
      this.currentUser.jobTitle,
    ]
      .filter((prop): prop is string => typeof prop === "string" && prop !== "")
      .map((prop) => prop.toLowerCase());

    const targetAudiences = Array.isArray(rule.audiences)
      ? rule.audiences.map((audience) =>
          typeof audience === "string" ? audience.toLowerCase() : "",
        )
      : [];

    switch (rule.operation) {
      case "anyOf":
        return targetAudiences.some((audience) =>
          userProperties.includes(audience),
        );

      case "allOf":
        return targetAudiences.every((audience) =>
          userProperties.includes(audience),
        );

      case "noneOf":
        return !targetAudiences.some((audience) =>
          userProperties.includes(audience),
        );

      default:
        return false;
    }
  }

  private isCurrentUser(person: IPersonField): boolean {
    if (!this.currentUser) return false;

    return (
      person.id === this.currentUser.id ||
      (person.email &&
        this.currentUser.email &&
        person.email.toLowerCase() === this.currentUser.email.toLowerCase()) ||
      (typeof person.loginName === "string" &&
        person.loginName.includes(this.currentUser.id))
    );
  }

  private isUserInGroup(group: IPersonField): boolean {
    if (group.isGroup !== true) {
      return false;
    }

    if (!this.userGroupIds.length && !this.spGroupIds.length) {
      return false;
    }

    if (
      group.id &&
      !isNaN(Number(group.id)) &&
      this.spGroupIds.includes(Number(group.id))
    ) {
      return true;
    }

    if (group.id && this.userGroupIds.includes(group.id)) {
      return true;
    }

    return this.userGroups.some(
      (userGroup) =>
        typeof userGroup === "string" &&
        typeof group.displayName === "string" &&
        userGroup.toLowerCase() === group.displayName.toLowerCase(),
    );
  }

  public getCurrentUser(): IUser | null {
    return this.currentUser;
  }

  public getUserDismissedAlerts(): string[] {
    try {
      if (!this.currentUser) return [];
      return (
        this.storageService.getFromSessionStorage<string[]>("DismissedAlerts", {
          userSpecific: true,
        }) || []
      );
    } catch (error) {
      logger.error(
        "UserTargetingService",
        "Error getting dismissed alerts",
        error,
      );
      return [];
    }
  }

  public addUserDismissedAlert(alertId: string): void {
    try {
      if (!this.currentUser) return;

      const dismissedAlerts = this.getUserDismissedAlerts();

      if (!dismissedAlerts.includes(alertId)) {
        dismissedAlerts.push(alertId);
        this.storageService.saveToSessionStorage(
          "DismissedAlerts",
          dismissedAlerts,
          { userSpecific: true },
        );
      }
    } catch (error) {
      logger.error(
        "UserTargetingService",
        "Error saving dismissed alert",
        error,
      );
    }
  }

  public getUserHiddenAlerts(): string[] {
    try {
      if (!this.currentUser) return [];
      return (
        this.storageService.getFromLocalStorage<string[]>("HiddenAlerts", {
          userSpecific: true,
        }) || []
      );
    } catch (error) {
      logger.error(
        "UserTargetingService",
        "Error getting hidden alerts",
        error,
      );
      return [];
    }
  }

  public addUserHiddenAlert(alertId: string): void {
    try {
      if (!this.currentUser) return;

      const hiddenAlerts = this.getUserHiddenAlerts();

      if (!hiddenAlerts.includes(alertId)) {
        hiddenAlerts.push(alertId);
        this.storageService.saveToLocalStorage("HiddenAlerts", hiddenAlerts, {
          userSpecific: true,
        });
      }
    } catch (error) {
      logger.error("UserTargetingService", "Error saving hidden alert", error);
    }
  }
}

export default UserTargetingService;
