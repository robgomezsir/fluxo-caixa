using System;
using fluxo_caixa_crud.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace fluxo_caixa_crud.Services
{
    public class WebViewBridge
    {
        private readonly SheetRegistry _registry;
        private readonly TabularCrudService _tabular;
        private readonly MatrixCrudService _matrix;
        private readonly SummaryService _summary;
        private readonly ConfigService _config;

        public WebViewBridge()
        {
            _registry = new SheetRegistry();
            _tabular = new TabularCrudService(_registry);
            _matrix = new MatrixCrudService(_registry);
            _summary = new SummaryService(_registry, _tabular);
            _config = new ConfigService(_registry, _matrix);
        }

        public void EnsureInitialized()
        {
            _registry.Initialize();
            ExcelHelper.RunWithExcel(() => _config.EnsureLookupsSheet());
        }

        public string HandleMessage(string json)
        {
            ApiRequest req = null;
            try
            {
                req = JsonConvert.DeserializeObject<ApiRequest>(json);
                if (req == null || string.IsNullOrEmpty(req.Action))
                    return ApiJson.Serialize(ApiResponse.Fail("", "Requisição inválida."));

                object data = Dispatch(req);
                return ApiJson.Serialize(ApiResponse.Ok(req.Id, data));
            }
            catch (Exception ex)
            {
                return ApiJson.Serialize(ApiResponse.Fail(req?.Id ?? "", ex.Message));
            }
        }

        private object Dispatch(ApiRequest req)
        {
            string action = req.Action.ToLowerInvariant();
            string entity = req.Entity ?? "";
            JObject payload = req.Payload ?? new JObject();

            switch (action)
            {
                case "ping":
                    return new { ok = true, app = "Fluxo de Caixa CRUD", version = "1.0" };

                case "meta":
                    return _registry.GetMeta();

                case "download":
                    return _config.DownloadWorkbook();

                case "save":
                    return SaveWorkbook();

                case "show_excel":
                    return ShowExcelBase();

                case "list":
                    return List(entity, payload);

                case "get":
                    return Get(entity, payload);

                case "create":
                    if (IsConfigEntity(entity))
                        return _config.Create(entity, payload);
                    return _tabular.Create(entity, payload);

                case "update":
                    if (entity == "matrix" || entity.StartsWith("matrix_"))
                        return _matrix.Upsert(payload);
                    return _tabular.Update(entity, payload["id"]?.ToString(), payload);

                case "upsert":
                    return _matrix.Upsert(payload);

                case "delete":
                    if (IsConfigEntity(entity))
                    {
                        _config.Delete(entity, payload);
                        return new { deleted = true };
                    }
                    Delete(entity, payload);
                    return new { deleted = true };

                case "resumo":
                    return _summary.GetResumo();

                default:
                    throw new InvalidOperationException("Ação desconhecida: " + req.Action);
            }
        }

        private static object SaveWorkbook()
        {
            string path = ExcelHelper.SaveHostWorkbook();
            return new { saved = true, path };
        }

        private static object ShowExcelBase()
        {
            ExcelHelper.ShowApplication();
            return new { visible = true };
        }

        private static bool IsConfigEntity(string entity)
        {
            return entity != null && entity.StartsWith("config_", StringComparison.OrdinalIgnoreCase);
        }

        private object List(string entity, JObject payload)
        {
            if (entity == "config_responsaveis")
                return _config.ListResponsaveis();
            if (entity == "config_unidades")
                return _config.ListUnidades();
            if (entity == "config_rubricas")
                return _config.ListRubricas(payload["sheet"]?.ToString());

            if (entity == "matrix_columns")
            {
                string sheet = payload["sheet"]?.ToString();
                if (string.IsNullOrEmpty(sheet))
                    throw new InvalidOperationException("Informe payload.sheet.");
                return _matrix.GetColumnLayout(sheet);
            }

            if (entity == "matrix" || entity.StartsWith("matrix_"))
            {
                string sheet = payload["sheet"]?.ToString();
                if (string.IsNullOrEmpty(sheet) && entity.StartsWith("matrix_"))
                {
                    var cfg = _registry.GetByEntity(entity);
                    sheet = cfg.SheetName;
                }
                if (string.IsNullOrEmpty(sheet))
                    throw new InvalidOperationException("Informe payload.sheet para listagem matricial.");
                return _matrix.List(sheet, payload["rubrica"]?.ToString());
            }

            if (entity == "matrix_premissas_cal" || entity == "matrix_premissas_neg" || entity == "matrix_capex0403")
            {
                var cfg = _registry.GetByEntity(entity);
                return _matrix.List(cfg.SheetName, payload["rubrica"]?.ToString());
            }

            if (entity == "rubricas")
            {
                string sheet = payload["sheet"]?.ToString();
                if (string.IsNullOrEmpty(sheet)) throw new InvalidOperationException("Informe payload.sheet.");
                return _matrix.GetRubricas(sheet);
            }

            return _tabular.List(entity);
        }

        private object Get(string entity, JObject payload)
        {
            if (entity == "matrix" || entity.StartsWith("matrix_"))
            {
                return _matrix.Get(
                    payload["sheet"]?.ToString(),
                    payload["rubrica"]?.ToString(),
                    payload["mes"]?.ToString(),
                    payload["tipo"]?.ToString());
            }
            return _tabular.Get(entity, payload["id"]?.ToString());
        }

        private void Delete(string entity, JObject payload)
        {
            if (entity == "matrix" || entity.StartsWith("matrix_"))
                _matrix.Delete(payload);
            else
                _tabular.Delete(entity, payload["id"]?.ToString());
        }
    }
}
