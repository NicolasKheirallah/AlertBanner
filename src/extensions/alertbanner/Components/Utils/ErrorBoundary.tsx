import * as React from 'react';
import { logger } from '../Services/LoggerService';
import { Text, Button, MessageBar } from '@fluentui/react-components';
import { ErrorCircle24Regular, ArrowClockwise24Regular } from '@fluentui/react-icons';
import styles from './ErrorBoundary.module.scss';

interface IErrorBoundaryState {
  hasError: boolean;
  error?: Error;
  errorInfo?: React.ErrorInfo;
  errorId?: string;
}

interface IErrorBoundaryProps {
  children: React.ReactNode;
  componentName?: string;
  fallback?: React.ComponentType<{ error: Error; reset: () => void }>;
  onError?: (error: Error, errorInfo: React.ErrorInfo) => void;
}

export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  private retryCount: number = 0;
  private maxRetries: number = 3;

  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false
    };
  }

  static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return {
      hasError: true,
      error,
      errorId: `error-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`
    };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    const componentName = this.props.componentName || 'Unknown Component';

    logger.error(componentName, 'React component error boundary caught an error', error, {
      errorInfo: {
        componentStack: errorInfo.componentStack,
        errorBoundary: componentName,
        retryCount: this.retryCount,
        timestamp: new Date().toISOString()
      },
      props: this.sanitizeProps(this.props),
      state: this.state
    });

    if (this.props.onError) {
      try {
        this.props.onError(error, errorInfo);
      } catch (handlerError) {
        logger.error(componentName, 'Error in custom error handler', handlerError);
      }
    }

    this.setState({
      error,
      errorInfo,
      errorId: `error-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`
    });
  }

  private sanitizeProps(props: IErrorBoundaryProps): any {
    const { children, onError, ...safeProps } = props;
    return {
      ...safeProps,
      hasChildren: !!children,
      hasOnError: !!onError
    };
  }

  private handleRetry = (): void => {
    if (this.retryCount < this.maxRetries) {
      this.retryCount++;
      const componentName = this.props.componentName || 'Unknown Component';
      
      logger.info(componentName, `Retrying component render (attempt ${this.retryCount}/${this.maxRetries})`);
      
      this.setState({
        hasError: false
      });
    } else {
      logger.warn(this.props.componentName || 'Unknown Component', 'Maximum retry attempts reached');
    }
  };

  private resetRetryCount(): void {
    if (this.retryCount > 0) {
      this.retryCount = 0;
    }
  }

  private handleCopyErrorDetails = async (): Promise<void> => {
    try {
      const errorDetails = {
        errorId: this.state.errorId,
        timestamp: new Date().toISOString(),
        component: this.props.componentName || 'Unknown Component',
        message: this.state.error?.message,
        stack: this.state.error?.stack,
        componentStack: this.state.errorInfo?.componentStack,
        userAgent: navigator.userAgent,
        url: window.location.href
      };

      await navigator.clipboard.writeText(JSON.stringify(errorDetails, null, 2));
      logger.info('ErrorBoundary', 'Error details copied to clipboard');
    } catch (clipboardError) {
      logger.warn('ErrorBoundary', 'Failed to copy error details to clipboard', clipboardError);
    }
  };

  render(): React.ReactNode {
    if (this.state.hasError) {
      if (this.props.fallback) {
        const FallbackComponent = this.props.fallback;
        return (
          <FallbackComponent 
            error={this.state.error!} 
            reset={this.handleRetry} 
          />
        );
      }

      const canRetry = this.retryCount < this.maxRetries;
      const componentName = this.props.componentName || 'Component';

      return (
        <div className={styles.errorContainer}>
          <MessageBar intent="error">
            <div className={styles.errorHeader}>
              <ErrorCircle24Regular />
              <Text weight="semibold">
                Something went wrong in {componentName}
              </Text>
            </div>
          </MessageBar>

          <div className={styles.errorMessage}>
            <Text size={300} className={styles.errorMessageText}>
              {this.state.error?.message || 'An unexpected error occurred'}
            </Text>
          </div>

          <div className={styles.errorId}>
            <Text size={200} className={styles.errorIdText}>
              Error ID: {this.state.errorId}
            </Text>
          </div>

          <div className={styles.errorActions}>
            {canRetry && (
              <Button 
                appearance="primary"
                icon={<ArrowClockwise24Regular />}
                onClick={this.handleRetry}
              >
                Try Again ({this.maxRetries - this.retryCount} attempts left)
              </Button>
            )}

            <Button 
              appearance="secondary"
              onClick={this.handleCopyErrorDetails}
            >
              Copy Error Details
            </Button>

            <Button 
              appearance="secondary"
              onClick={() => window.location.reload()}
            >
              Reload Page
            </Button>
          </div>

          {process.env.NODE_ENV === 'development' && (
            <details className={styles.errorDetails}>
              <summary className={styles.errorDetailsSummary}>
                Development Error Details
              </summary>
              <pre className={styles.errorDetailsCode}>
                {this.state.error?.stack}
              </pre>
              {this.state.errorInfo?.componentStack && (
                <pre className={styles.errorDetailsCode}>
                  {this.state.errorInfo.componentStack}
                </pre>
              )}
            </details>
          )}
        </div>
      );
    }

    this.resetRetryCount();
    return this.props.children;
  }
}

export function withErrorBoundary<P extends object>(
  Component: React.ComponentType<P>,
  errorBoundaryProps?: Omit<IErrorBoundaryProps, 'children'>
) {
  const ComponentWithErrorBoundary = (props: P) => (
    <ErrorBoundary {...errorBoundaryProps}>
      <Component {...props} />
    </ErrorBoundary>
  );

  ComponentWithErrorBoundary.displayName = `withErrorBoundary(${Component.displayName || Component.name})`;

  return ComponentWithErrorBoundary;
}

export default ErrorBoundary;