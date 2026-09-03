"use client";
import { zodResolver } from "@hookform/resolvers/zod";
import { Controller, useForm } from "react-hook-form";
import { toast } from "sonner";
import * as z from "zod";

import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Field,
  FieldError,
  FieldGroup,
  FieldLabel,
} from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { processarExcel } from "./_actions";
import { useState } from "react";
import {
  Attachment,
  AttachmentAction,
  AttachmentActions,
  AttachmentContent,
  AttachmentDescription,
  AttachmentMedia,
  AttachmentTitle,
} from "@/components/ui/attachment";
import { FileIcon } from "lucide-react";

const formSchema = z.object({
  nomeAbaOriginal: z.string().min(1, "Campo obrigatório"),
  colunas: z.string().min(1, "Campo obrigatório"),
  nomeAbaDestino: z.string().min(1, "Campo obrigatório"),
  planilha: z
    .file()
    .optional()
    .refine((file) => file !== undefined, {
      message: "Campo obrigatório",
    })
    .refine((file) => !file || file.size > 0, {
      message: "O arquivo não pode estar vazio",
    })
    .refine(
      (file) =>
        !file ||
        [
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "application/vnd.ms-excel",
        ].includes(file.type),
      {
        message: "Apenas arquivos Excel são permitidos",
      },
    ),
});

const MesclarPlanilhas = () => {
  const [loading, setLoading] = useState(false);
  const form = useForm<z.infer<typeof formSchema>>({
    resolver: zodResolver(formSchema),
    defaultValues: {
      nomeAbaOriginal: "",
      colunas: "",
      nomeAbaDestino: "",
      planilha: undefined,
    },
  });

  const onSubmit = async (data: z.infer<typeof formSchema>) => {
    try {
      setLoading(true);
      const colunasArray = data.colunas
        .split(",")
        .map((coluna) => coluna.trim());

      const resultado = await processarExcel(
        data.planilha as File,
        data.nomeAbaOriginal,
        colunasArray,
        data.nomeAbaDestino,
      );

      const binaryString = atob(resultado.arquivo);

      const bytes = new Uint8Array(binaryString.length);

      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }

      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);

      const link = document.createElement("a");

      link.href = url;
      link.download = resultado.nome;

      document.body.appendChild(link);
      link.click();
      link.remove();

      URL.revokeObjectURL(url);

      toast.success("Planilha mesclada com sucesso!");
    } catch (error) {
      console.error(error);
      if (error instanceof Error) {
        toast.error(error.message);
      } else {
        toast.error(
          "Ocorreu um erro ao mesclar a planilha, por favor tente novamente mais tarde",
        );
      }
    }
    setLoading(false);
  };

  return (
    <Card className="w-full sm:max-w-md">
      <CardHeader>
        <CardTitle>Mesclar Planilhas</CardTitle>
        <CardDescription>
          Mescla planilhas de forma automática, você precisa selecionar o nome
          da aba original, as colunas que deseja mesclar e o nome da aba para
          destino.
        </CardDescription>
      </CardHeader>
      <CardContent>
        <form id="form-rhf-demo" onSubmit={form.handleSubmit(onSubmit)}>
          <FieldGroup>
            <Controller
              name="nomeAbaOriginal"
              control={form.control}
              render={({ field, fieldState }) => (
                <Field data-invalid={fieldState.invalid}>
                  <FieldLabel htmlFor="aba-original">
                    Nome da Aba Original
                  </FieldLabel>
                  <Input
                    {...field}
                    id="aba-original"
                    aria-invalid={fieldState.invalid}
                    placeholder="Digite o nome da aba original"
                    autoComplete="off"
                  />
                  {fieldState.invalid && (
                    <FieldError errors={[fieldState.error]} />
                  )}
                </Field>
              )}
            />
            <Controller
              name="colunas"
              control={form.control}
              render={({ field, fieldState }) => (
                <Field data-invalid={fieldState.invalid}>
                  <FieldLabel htmlFor="colunas">Colunas</FieldLabel>
                  <Input
                    {...field}
                    id="colunas"
                    aria-invalid={fieldState.invalid}
                    placeholder="Digite as colunas separadas por vírgula, Ex: A, B, C"
                    autoComplete="off"
                  />
                  {fieldState.invalid && (
                    <FieldError errors={[fieldState.error]} />
                  )}
                </Field>
              )}
            />
            <Controller
              name="nomeAbaDestino"
              control={form.control}
              render={({ field, fieldState }) => (
                <Field data-invalid={fieldState.invalid}>
                  <FieldLabel htmlFor="aba-destino">
                    Nome da Aba Destino
                  </FieldLabel>
                  <Input
                    {...field}
                    id="aba-destino"
                    aria-invalid={fieldState.invalid}
                    placeholder="Digite o nome da aba destino"
                    autoComplete="off"
                  />
                  {fieldState.invalid && (
                    <FieldError errors={[fieldState.error]} />
                  )}
                </Field>
              )}
            />
            <Controller
              name="planilha"
              control={form.control}
              render={({ field, fieldState }) => (
                <Field data-invalid={fieldState.invalid}>
                  <FieldLabel htmlFor="planilha">Planilha</FieldLabel>
                  {field.value ? (
                    <Attachment className="w-full">
                      <AttachmentMedia>
                        <FileIcon />
                      </AttachmentMedia>

                      <AttachmentContent>
                        <AttachmentTitle>{field.value.name}</AttachmentTitle>

                        <AttachmentDescription>
                          {(field.value.size / 1024).toFixed(1)} KB
                        </AttachmentDescription>
                      </AttachmentContent>

                      <AttachmentActions>
                        <AttachmentAction
                          type="button"
                          onClick={() => {
                            field.onChange(undefined);
                          }}
                          aria-label="Remover arquivo"
                        >
                          ✕
                        </AttachmentAction>
                      </AttachmentActions>
                    </Attachment>
                  ) : (
                    <Input
                      id="planilha"
                      type="file"
                      accept=".xlsx,.xls"
                      name={field.name}
                      ref={field.ref}
                      onBlur={field.onBlur}
                      onChange={(event) => {
                        const file = event.target.files?.[0];
                        field.onChange(file);
                      }}
                      aria-invalid={fieldState.invalid}
                    />
                  )}
                  {fieldState.invalid && (
                    <FieldError errors={[fieldState.error]} />
                  )}
                </Field>
              )}
            />
          </FieldGroup>
        </form>
      </CardContent>
      <CardFooter>
        <Field orientation="horizontal">
          <Button type="button" variant="outline" onClick={() => form.reset()}>
            Limpar
          </Button>
          <Button loading={loading} type="submit" form="form-rhf-demo">
            Mesclar
          </Button>
        </Field>
      </CardFooter>
    </Card>
  );
};

export default MesclarPlanilhas;
