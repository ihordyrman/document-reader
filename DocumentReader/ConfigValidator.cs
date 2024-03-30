using FluentValidation;

namespace DocumentReader;

public class ConfigValidator : AbstractValidator<Config>
{
    public ConfigValidator()
        => RuleFor(config => config.Areas)
            .NotEmpty()
            .WithMessage("Areas must not be empty.")
            .ForEach(
                area =>
                {
                    area.NotNull().WithMessage("Area must not be null.");
                    area.ChildRules(
                        areaRules =>
                        {
                            areaRules.RuleFor(a => a.Name).NotEmpty().WithMessage("Area name must not be empty.");
                            areaRules.RuleFor(a => a.X)
                                .GreaterThanOrEqualTo(0)
                                .WithMessage(x => $"Area X must be greater than or equal to 0. Section: {x.Name}");
                            areaRules.RuleFor(a => a.Y)
                                .GreaterThanOrEqualTo(0)
                                .WithMessage(x => $"Area Y must be greater than or equal to 0. Section: {x.Name}");
                            areaRules.RuleFor(a => a.Width)
                                .GreaterThan(0)
                                .WithMessage(x => $"Area width must be greater than 0. {x.Name}");
                            areaRules.RuleFor(a => a.Height)
                                .GreaterThan(0)
                                .WithMessage(x => $"Area height must be greater than 0. {x.Name}");
                        });
                });
}
