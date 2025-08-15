import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import { z } from "zod";
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import { validation } from "../../../../utils/validation.js";
import request, { CliRequestOptions } from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { zod } from "../../../../utils/zod.js";

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()),
    displayName: zod.alias('n', z.string().optional()),
    properties: zod.alias('p', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_GET;
  }

  public get description(): string {
    return 'Gets information about a specific administrative unit';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.id !== !options.displayName, {
        message: 'Specify either id or displayName, but not both'
      })
      .refine(options => options.id || options.displayName, {
        message: 'Specify either id or displayName'
      });
  }

  constructor() {
    super();

    this.#initTelemetry();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnit: AdministrativeUnit;

    try {
      if (args.options.id) {
        administrativeUnit = await this.getAdministrativeUnitById(args.options.id, args.options.properties);
      }
      else {
        administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName!);
      }

      await logger.log(administrativeUnit);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  async getAdministrativeUnitById(id: string, properties?: string): Promise<AdministrativeUnit> {
    const queryParameters: string[] = [];

    if (properties) {
      const allProperties = properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/administrativeUnits/${id}${queryString}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<AdministrativeUnit>(requestOptions);
  }
}

export default new EntraAdministrativeUnitGetCommand();