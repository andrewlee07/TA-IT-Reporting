export class PlatformError extends Error {
  status: number;

  constructor(message: string, status = 400) {
    super(message);
    this.name = "PlatformError";
    this.status = status;
  }
}

export class PlatformUnauthorizedError extends PlatformError {
  constructor(message = "Platform session is required.") {
    super(message, 401);
    this.name = "PlatformUnauthorizedError";
  }
}

export class PlatformForbiddenError extends PlatformError {
  constructor(message = "You do not have access to this platform resource.") {
    super(message, 403);
    this.name = "PlatformForbiddenError";
  }
}

export class PlatformNotFoundError extends PlatformError {
  constructor(message = "Platform resource not found.") {
    super(message, 404);
    this.name = "PlatformNotFoundError";
  }
}
